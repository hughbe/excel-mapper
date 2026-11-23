using System.Collections.Concurrent;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq.Expressions;
using System.Reflection;
#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
#endif

namespace ExcelMapper.Binding;

/// <summary>
/// Creates and caches compiled delegates that set members and call constructors, which are much
/// faster than reflection when called for every cell of every row.
/// </summary>
/// <remarks>
/// Delegates are cached per member or constructor, so maps that are created for each importer (such
/// as auto-mapped types) only pay the cost of compiling once. Where code can't be compiled, for
/// example when running with Native AOT, reflection is used instead.
/// </remarks>
internal static class CompiledAccessors
{
    private static readonly ConcurrentDictionary<ConstructorInfo, Func<object?[], object>> s_constructors = new();
    private static readonly ConcurrentDictionary<Type, Func<object>> s_parameterlessConstructors = new();
    private static readonly ConcurrentDictionary<(MemberInfo Member, Type OwnerType, Type ValueType), Delegate> s_setters = new();

    private static bool CanCompile
    {
        get
        {
#if NET8_0_OR_GREATER
            return RuntimeFeature.IsDynamicCodeCompiled;
#else
            return true;
#endif
        }
    }

    /// <summary>
    /// Gets a delegate that sets a property or field of a <typeparamref name="TOwner"/> to a
    /// <typeparamref name="TValue"/>, without boxing either: a struct is set in place, through the
    /// reference. If <typeparamref name="TValue"/> isn't the member's type, the value is converted to
    /// it, and null sets a value type to its default value. A setter an
    /// <see cref="ExcelMappingContext"/> registered is used if there is one.
    /// </summary>
    // Keep in sync with the source generator's TypeWalker.AddSetter, which registers a setter for
    // each member set here. See "Auto-mapping and the source generator" in CLAUDE.md.
    public static ExcelMemberSetter<TOwner, TValue> GetSetter<TOwner, TValue>(MemberInfo member)
        => (ExcelMemberSetter<TOwner, TValue>)s_setters.GetOrAdd((member, typeof(TOwner), typeof(TValue)), static key => CreateSetter<TOwner, TValue>(key.Member));

    private static ExcelMemberSetter<TOwner, TValue> CreateSetter<TOwner, TValue>(MemberInfo member)
    {
        // A setter a mapping context registered: ordinary code, written by the source generator.
        if (RegisteredAccessors.GetSetter<TOwner, TValue>(member.Name) is { } registered)
        {
            return registered;
        }

        var memberType = member.MemberType();

        // A read-only property is only mapped when it's a constructor parameter (see
        // ConstructionPlan<T>.Create), so it's never set.
        Debug.Assert(member is not PropertyInfo { CanWrite: false }, $"Property \"{member.Name}\" is read-only.");

        // Readonly fields can't be assigned in an expression, and code can't be compiled with
        // native AOT: use reflection, which boxes the value, and a struct instance.
        if (!CanCompile || member is FieldInfo { IsInitOnly: true })
        {
            return CreateReflectionSetter<TOwner, TValue>(member);
        }

        var instance = Expression.Parameter(typeof(TOwner).MakeByRefType(), "instance");
        var value = Expression.Parameter(typeof(TValue), "value");
        var declaringType = member.DeclaringType!;
        Expression typedInstance = declaringType == typeof(TOwner) ? instance : Expression.Convert(instance, declaringType);
        Expression typedValue = typeof(TValue) == memberType
            ? value
            : ConvertValue(Expression.Convert(value, typeof(object)), memberType);

        Expression body = member is PropertyInfo property
            ? Expression.Call(typedInstance, property.SetMethod!, typedValue)
            : Expression.Assign(Expression.Field(typedInstance, (FieldInfo)member), typedValue);

        return Expression.Lambda<ExcelMemberSetter<TOwner, TValue>>(body, instance, value).Compile();
    }

    private static ExcelMemberSetter<TOwner, TValue> CreateReflectionSetter<TOwner, TValue>(MemberInfo member)
    {
        Action<object?, object?> setValue = member is PropertyInfo property ? property.SetValue : ((FieldInfo)member).SetValue;

        if (typeof(TOwner).IsValueType)
        {
            // Set the struct in a box, then copy it back.
            return (ref TOwner instance, TValue value) =>
            {
                object boxed = instance!;
                setValue(boxed, value);
                instance = (TOwner)boxed;
            };
        }

        return (ref TOwner instance, TValue value) => setValue(instance, value);
    }

    /// <summary>
    /// Gets a delegate that calls a constructor with one argument, without allocating an argument array.
    /// </summary>
    public static Func<object, object> GetConstructor1(ConstructorInfo constructor)
    {
        if (!CanCompile)
        {
            return argument => constructor.Invoke([argument]);
        }

        var argument = Expression.Parameter(typeof(object), "argument");
        var parameterType = constructor.GetParameters()[0].ParameterType;
        var body = Expression.Convert(Expression.New(constructor, Expression.Convert(argument, parameterType)), typeof(object));
        return Expression.Lambda<Func<object, object>>(body, argument).Compile();
    }

    /// <summary>
    /// Gets a delegate that calls an instance method with one argument, such as a collection's
    /// <c>Add(T)</c>, without boxing value types or allocating an argument array. A struct instance
    /// is changed inside its box.
    /// </summary>
    public static Action<object, T> GetAction<T>(MethodInfo method)
    {
        if (!CanCompile)
        {
            return (instance, value) => method.InvokeUnwrapped(instance, [value]);
        }

        var instance = Expression.Parameter(typeof(object), "instance");
        var value = Expression.Parameter(typeof(T), "value");
        var body = Expression.Call(GetBoxedInstance(instance, method.DeclaringType!), method, ConvertArgument(value, method.GetParameters()[0].ParameterType));
        return Expression.Lambda<Action<object, T>>(body, instance, value).Compile();
    }

    /// <summary>
    /// Gets a delegate that calls an instance method with two arguments, such as a dictionary's
    /// <c>Add(TKey, TValue)</c>. A struct instance is changed inside its box.
    /// </summary>
    public static Action<object, T1, T2> GetAction<T1, T2>(MethodInfo method)
    {
        if (!CanCompile)
        {
            return (instance, value1, value2) => method.InvokeUnwrapped(instance, [value1, value2]);
        }

        var instance = Expression.Parameter(typeof(object), "instance");
        var value1 = Expression.Parameter(typeof(T1), "value1");
        var value2 = Expression.Parameter(typeof(T2), "value2");
        var parameters = method.GetParameters();
        var body = Expression.Call(
            GetBoxedInstance(instance, method.DeclaringType!),
            method,
            ConvertArgument(value1, parameters[0].ParameterType),
            ConvertArgument(value2, parameters[1].ParameterType));
        return Expression.Lambda<Action<object, T1, T2>>(body, instance, value1, value2).Compile();
    }

    // The instance a method is called on: a struct is unboxed in place, so the method changes the
    // boxed struct rather than a copy.
    private static Expression GetBoxedInstance(ParameterExpression instance, Type declaringType)
        => declaringType.IsValueType ? Expression.Unbox(instance, declaringType) : Expression.Convert(instance, declaringType);

    private static Expression ConvertArgument(Expression value, Type parameterType)
        => value.Type == parameterType ? value : Expression.Convert(value, parameterType);

    /// <summary>
    /// Gets a delegate that calls a constructor with an array of arguments. Null arguments for value
    /// type parameters pass the default value, matching <see cref="ConstructorInfo.Invoke(object[])"/>.
    /// </summary>
    public static Func<object?[], object> GetConstructor(ConstructorInfo constructor)
        => s_constructors.GetOrAdd(constructor, static c => CreateConstructor(c));

    private static Func<object?[], object> CreateConstructor(ConstructorInfo constructor)
    {
        if (!CanCompile)
        {
            return constructor.Invoke;
        }

        var arguments = Expression.Parameter(typeof(object?[]), "arguments");
        var parameters = constructor.GetParameters();
        var argumentExpressions = new Expression[parameters.Length];
        for (int i = 0; i < parameters.Length; i++)
        {
            argumentExpressions[i] = ConvertValue(Expression.ArrayIndex(arguments, Expression.Constant(i)), parameters[i].ParameterType);
        }

        var body = Expression.Convert(Expression.New(constructor, argumentExpressions), typeof(object));
        return Expression.Lambda<Func<object?[], object>>(body, arguments).Compile();
    }

    // The compiled constructors of each type, by the arguments they take from bound members.
    private static class BoundConstructors<TOwner>
    {
        public static readonly ConcurrentDictionary<BoundConstructorKey, Func<BoundMember<TOwner>[], TOwner>> Cache = new();

        public static readonly MethodInfo GetValueOrDefault = new Func<BoundMember<TOwner>, object?, object?>(BoundMember<TOwner>.GetValueOrDefault).Method;
    }

    /// <summary>
    /// Gets a delegate that calls a constructor of <typeparamref name="TOwner"/> with the values
    /// bound members read for the row (see <see cref="BoundMember{TOwner}.ReadValue"/>), without an
    /// argument array, without boxing the values of members that keep them typed, and without
    /// boxing the instance. Returns null where code can't be compiled.
    /// </summary>
    /// <param name="constructor">The constructor to call.</param>
    /// <param name="arguments">For each parameter, the member whose value it takes, if any, and its
    /// default value for when there is no value.</param>
    public static Func<BoundMember<TOwner>[], TOwner>? GetConstructor<TOwner>(ConstructorInfo constructor, BoundConstructorArgument[] arguments)
    {
        if (!CanCompile)
        {
            return null;
        }

        var key = new BoundConstructorKey(constructor, arguments);
        var cache = BoundConstructors<TOwner>.Cache;
        if (!cache.TryGetValue(key, out var result))
        {
            result = cache.GetOrAdd(key, CreateConstructor<TOwner>(constructor, arguments));
        }

        return result;
    }

    private static Func<BoundMember<TOwner>[], TOwner> CreateConstructor<TOwner>(ConstructorInfo constructor, BoundConstructorArgument[] arguments)
    {
        var members = Expression.Parameter(typeof(BoundMember<TOwner>[]), "members");
        var parameters = constructor.GetParameters();
        var argumentExpressions = new Expression[parameters.Length];
        var variables = new List<ParameterExpression>();
        for (int i = 0; i < parameters.Length; i++)
        {
            var parameterType = parameters[i].ParameterType;
            var argument = arguments[i];
            var defaultValue = Expression.Constant(argument.DefaultValue, typeof(object));
            if (argument.MemberIndex < 0)
            {
                argumentExpressions[i] = ConvertValue(defaultValue, parameterType);
                continue;
            }

            var member = Expression.ArrayIndex(members, Expression.Constant(argument.MemberIndex));
            if (argument.TypedMethod is { } typedMethod)
            {
                argumentExpressions[i] = Expression.Call(
                    typedMethod,
                    Expression.Convert(member, typedMethod.GetParameters()[0].ParameterType),
                    ConvertValue(defaultValue, parameterType));
            }
            else
            {
                // Read the value once, as converting it checks it for null first.
                var value = Expression.Variable(typeof(object));
                variables.Add(value);
                argumentExpressions[i] = Expression.Block(
                    Expression.Assign(value, Expression.Call(BoundConstructors<TOwner>.GetValueOrDefault, member, defaultValue)),
                    ConvertValue(value, parameterType));
            }
        }

        var body = Expression.New(constructor, argumentExpressions);
        return Expression.Lambda<Func<BoundMember<TOwner>[], TOwner>>(Expression.Block(variables, body), members).Compile();
    }

    /// <summary>
    /// Gets a delegate that creates a <typeparamref name="T"/>, a struct or a class with a public
    /// parameterless constructor, without boxing it.
    /// </summary>
    /// <remarks>
    /// Compiled once per type: compiling takes tens of microseconds, and every importer that reads
    /// a type asks for its constructor.
    /// </remarks>
    public static Func<T> GetParameterlessConstructor<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] T>()
        => ParameterlessConstructor<T>.Instance;

    private static class ParameterlessConstructor<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] T>
    {
        public static readonly Func<T> Instance = CanCompile ? Expression.Lambda<Func<T>>(Expression.New(typeof(T))).Compile() : Activator.CreateInstance<T>;
    }

    /// <summary>
    /// Gets a delegate that creates an instance of a type with a public parameterless constructor, or
    /// a struct.
    /// </summary>
    public static Func<object> GetParameterlessConstructor([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] Type type)
    {
        if (!s_parameterlessConstructors.TryGetValue(type, out var constructor))
        {
            // Not a GetOrAdd lambda: the trimming annotation on the type doesn't flow into lambdas.
            constructor = s_parameterlessConstructors.GetOrAdd(type, CreateParameterlessConstructor(type));
        }

        return constructor;
    }

    private static Func<object> CreateParameterlessConstructor([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] Type type)
    {
        if (!CanCompile || type.IsAbstract || type.IsInterface || (!type.IsValueType && type.GetConstructor(Type.EmptyTypes) is null))
        {
            // Let Activator report the error for types that can't be created.
            return () => Activator.CreateInstance(type)!;
        }

        var body = Expression.Convert(Expression.New(type), typeof(object));
        return Expression.Lambda<Func<object>>(body).Compile();
    }

    private static Expression ConvertValue(Expression value, Type type)
    {
        var converted = Expression.Convert(value, type);
        if (type.IsValueType && Nullable.GetUnderlyingType(type) is null)
        {
            // Unboxing null would throw, so use the default value instead.
            return Expression.Condition(
                Expression.Equal(value, Expression.Constant(null)),
                Expression.Default(type),
                converted);
        }

        return converted;
    }
}

/// <summary>
/// How a constructor compiled by <see cref="CompiledAccessors.GetConstructor{TOwner}(ConstructorInfo, BoundConstructorArgument[])"/>
/// gets a parameter's value.
/// </summary>
/// <param name="memberIndex">The index of the bound member whose value the parameter takes, or -1
/// to always pass <paramref name="defaultValue"/>.</param>
/// <param name="typedMethod">The method that returns the member's value without boxing (see
/// <see cref="IBindableMap.GetArgumentMethod{TOwner}(Type)"/>), or null to read it as an object.</param>
/// <param name="defaultValue">The value passed when the member has no value. Null passes the
/// parameter type's default value.</param>
internal readonly struct BoundConstructorArgument(int memberIndex, MethodInfo? typedMethod, object? defaultValue)
{
    public int MemberIndex { get; } = memberIndex;

    public MethodInfo? TypedMethod { get; } = typedMethod;

    public object? DefaultValue { get; } = defaultValue;
}

/// <summary>
/// The key for a compiled constructor over bound members: the delegate depends on the constructor
/// and, for each parameter, which member it takes and whether that member's value is typed. The
/// default values come from the constructor.
/// </summary>
internal sealed class BoundConstructorKey(ConstructorInfo constructor, BoundConstructorArgument[] arguments) : IEquatable<BoundConstructorKey>
{
    public bool Equals(BoundConstructorKey? other)
    {
        if (other is null || other.Constructor() != constructor || other.Arguments().Length != arguments.Length)
        {
            return false;
        }

        var otherArguments = other.Arguments();
        for (int i = 0; i < arguments.Length; i++)
        {
            if (arguments[i].MemberIndex != otherArguments[i].MemberIndex ||
                (arguments[i].TypedMethod is null) != (otherArguments[i].TypedMethod is null))
            {
                return false;
            }
        }

        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as BoundConstructorKey);

    public override int GetHashCode()
    {
        var hash = constructor.GetHashCode();
        foreach (var argument in arguments)
        {
            hash = (hash * 31) + (argument.MemberIndex * 2) + (argument.TypedMethod is null ? 0 : 1);
        }

        return hash;
    }

    private ConstructorInfo Constructor() => constructor;

    private BoundConstructorArgument[] Arguments() => arguments;
}
