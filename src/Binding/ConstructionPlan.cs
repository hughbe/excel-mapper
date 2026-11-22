using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Binding;

/// <summary>
/// How a class map creates its type, decided once when the first row is read: with a
/// <b>parameterless constructor</b>, after which every member is set; or, for records and
/// classes with read-only properties, with the public <b>constructor</b> whose parameters match
/// mapped members by name and type (or have default values), after which the members that
/// aren't parameters are set.
/// </summary>
internal sealed class ConstructionPlan<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>
{
    private ConstructionPlan(Func<T> createInstance)
    {
        CreateInstance = createInstance;
        MemberParameterIndices = [];
        DefaultArguments = [];
    }

    private ConstructionPlan(ConstructorInfo constructor, int[] memberParameterIndices, object?[] defaultArguments, Func<BoundMember<T>[], T>? createFromMembers)
    {
        Constructor = constructor;
        MemberParameterIndices = memberParameterIndices;
        DefaultArguments = defaultArguments;
        CreateFromMembers = createFromMembers;
        InvokeConstructor = createFromMembers is null ? CompiledAccessors.GetConstructor(constructor) : null;
    }

    /// <summary>
    /// Gets the constructor to call, or null to use the parameterless constructor.
    /// </summary>
    public ConstructorInfo? Constructor { get; }

    /// <summary>
    /// Gets, for each member, the index of the constructor parameter it is passed to, or -1.
    /// </summary>
    private int[] MemberParameterIndices { get; }

    // The arguments of parameters no member is passed to: their default values.
    private object?[] DefaultArguments { get; }

    /// <summary>
    /// Gets a delegate that creates an instance with the parameterless constructor, if
    /// <see cref="Constructor"/> is null.
    /// </summary>
    public Func<T>? CreateInstance { get; }

    // Calls the constructor with the values the bound members read: a constructor a mapping context
    // registered, or a compiled one; null where code can't be compiled.
    private Func<BoundMember<T>[], T>? CreateFromMembers { get; }

    // Calls the constructor with an array of arguments, where code can't be compiled.
    private Func<object?[], object>? InvokeConstructor { get; }

    /// <summary>
    /// Calls <see cref="Constructor"/> with the values the bound members read (see
    /// <see cref="BoundMember{TOwner}.ReadValue"/>).
    /// </summary>
    public T Construct(BoundMember<T>[] members)
        => CreateFromMembers is { } createFromMembers
            ? createFromMembers(members)
            : (T)InvokeConstructor!(GetArguments(members));

    /// <summary>
    /// Sets the members that weren't passed to the constructor.
    /// </summary>
    public void SetOtherMembers(ref T instance, BoundMember<T>[] members)
    {
        for (int i = 0; i < members.Length; i++)
        {
            if (MemberParameterIndices[i] < 0 && members[i].HasValue)
            {
                members[i].SetValue(ref instance);
            }
        }
    }

    // Where constructors can't be compiled: the constructor's arguments as an array.
    private object?[] GetArguments(BoundMember<T>[] members)
    {
        var arguments = (object?[])DefaultArguments.Clone();
        for (int i = 0; i < members.Length; i++)
        {
            var parameterIndex = MemberParameterIndices[i];
            if (parameterIndex >= 0 && members[i].HasValue)
            {
                arguments[parameterIndex] = members[i].Value;
            }
        }

        return arguments;
    }

    /// <summary>
    /// Decides how to create a <typeparamref name="T"/> with the given members.
    /// </summary>
    /// <exception cref="ExcelMappingException">The type can't be created with these members.</exception>
    public static ConstructionPlan<T> Create(IReadOnlyList<ClassMapMember> classMembers)
    {
        var type = typeof(T);
        var hasReadOnlyMember = false;
        foreach (var member in classMembers)
        {
            if (member.IsReadOnly)
            {
                hasReadOnlyMember = true;
                break;
            }
        }

        // Prefer a parameterless constructor and setting members, unless a member can only be
        // set through a constructor.
        if (!hasReadOnlyMember && ConstructorUtilities.HasParameterlessConstructor(type))
        {
            return new ConstructionPlan<T>(CompiledAccessors.GetParameterlessConstructor<T>());
        }

        if (type.IsInterface || type.IsAbstract)
        {
            throw new ExcelMappingException($"Cannot create an instance of type \"{type}\" as it is an interface or abstract class.");
        }

        var members = new MemberInfo[classMembers.Count];
        for (int i = 0; i < members.Length; i++)
        {
            members[i] = classMembers[i].Member;
        }

        var constructor = ConstructorUtilities.FindConstructor(type, members, out var parameterMemberIndices);
        if (constructor is null)
        {
            throw new ExcelMappingException(
                $"Cannot create an instance of type \"{type}\". It needs a public parameterless constructor, or a public " +
                "constructor whose parameters each match a mapped property or field (by name and type) or have a default value.");
        }

        var memberParameterIndices = new int[members.Length];
        for (int i = 0; i < memberParameterIndices.Length; i++)
        {
            memberParameterIndices[i] = -1;
        }

        var parameters = constructor.GetParameters();
        var defaultArguments = new object?[parameters.Length];
        for (int i = 0; i < parameters.Length; i++)
        {
            if (parameterMemberIndices[i] >= 0)
            {
                memberParameterIndices[parameterMemberIndices[i]] = i;
            }

            // Reflection passes a zero-initialized value for null value type arguments.
            defaultArguments[i] = parameters[i].HasDefaultValue ? parameters[i].DefaultValue : null;
        }

        for (int i = 0; i < members.Length; i++)
        {
            if (memberParameterIndices[i] == -1 && classMembers[i].IsReadOnly)
            {
                throw new ExcelMappingException($"Property \"{members[i].Name}\" of type \"{type}\" is read-only and is not a parameter of the constructor {constructor}.");
            }
        }

        var boundArguments = new BoundConstructorArgument[parameters.Length];
        for (int i = 0; i < parameters.Length; i++)
        {
            var memberIndex = parameterMemberIndices[i];
            var typedMethod = memberIndex >= 0 && classMembers[memberIndex].Map is IBindableMap bindableMap
                ? bindableMap.GetArgumentMethod<T>(parameters[i].ParameterType)
                : null;
            boundArguments[i] = new BoundConstructorArgument(memberIndex, typedMethod, defaultArguments[i]);
        }

        return new ConstructionPlan<T>(constructor, memberParameterIndices, defaultArguments, GetRegisteredConstructor(parameters, boundArguments) ?? CompiledAccessors.GetConstructor<T>(constructor, boundArguments));
    }

    // A constructor a mapping context registered: ordinary code, written by the source generator,
    // which reads its arguments typed.
    private static Func<BoundMember<T>[], T>? GetRegisteredConstructor(ParameterInfo[] parameters, BoundConstructorArgument[] boundArguments)
    {
        var parameterTypes = new Type[parameters.Length];
        for (int i = 0; i < parameters.Length; i++)
        {
            parameterTypes[i] = parameters[i].ParameterType;
        }

        return RegisteredAccessors.GetConstructor<T>(parameterTypes) is { } constructor
            ? members => constructor(new ExcelConstructorArguments<T>(members, boundArguments))
            : null;
    }
}
