using System.Diagnostics.CodeAnalysis;
using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace ExcelMapper.Utilities;

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
    private static readonly ConcurrentDictionary<MemberInfo, MemberSetValueDelegate> s_setters = new();
    private static readonly ConcurrentDictionary<ConstructorInfo, Func<object?[], object>> s_constructors = new();
    private static readonly ConcurrentDictionary<Type, Func<object>> s_parameterlessConstructors = new();

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
    /// Gets a delegate that sets a writable property or field. A null value sets the member to its
    /// default value, matching <see cref="PropertyInfo.SetValue(object, object)"/>.
    /// </summary>
    public static MemberSetValueDelegate GetSetter(MemberInfo member)
        => s_setters.GetOrAdd(member, static m => CreateSetter(m));

    private static MemberSetValueDelegate CreateSetter(MemberInfo member)
    {
        var declaringType = member.DeclaringType!;

        // Setting a member of a boxed struct through an expression would modify a copy, and readonly
        // fields can't be assigned in an expression, so use reflection for these.
        var useReflection = !CanCompile || declaringType.IsValueType || member is FieldInfo { IsInitOnly: true };
        if (useReflection)
        {
            return member switch
            {
                PropertyInfo reflectedProperty => reflectedProperty.SetValue,
                FieldInfo reflectedField => reflectedField.SetValue,
                _ => throw new ArgumentException($"Member \"{member.Name}\" is not a field or property.", nameof(member))
            };
        }

        var instance = Expression.Parameter(typeof(object), "instance");
        var value = Expression.Parameter(typeof(object), "value");
        var typedInstance = Expression.Convert(instance, declaringType);

        Expression body;
        if (member is PropertyInfo property)
        {
            body = Expression.Call(typedInstance, property.SetMethod!, ConvertValue(value, property.PropertyType));
        }
        else
        {
            var field = (FieldInfo)member;
            body = Expression.Assign(Expression.Field(typedInstance, field), ConvertValue(value, field.FieldType));
        }

        return Expression.Lambda<MemberSetValueDelegate>(body, instance, value).Compile();
    }

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
