using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Mappers;

namespace ExcelMapper.AutoMapping;

/// <summary>
/// Creates the generic types auto-mapping needs for types found at run time, such as the
/// <see cref="TypeRule{T}"/> of a member's type or the <see cref="EnumMapper{TEnum}"/> of an enum.
/// They're made with <c>MakeGenericType</c>, which native AOT can't do for types it didn't compile,
/// unless an <see cref="ExcelMappingContext"/> registered them (see <see cref="ExcelTypeRegistrar"/>):
/// a registration creates the type with <c>new</c>, which native AOT compiles.
/// </summary>
internal static class GenericTypes
{
    private static readonly ConcurrentDictionary<(Type Definition, Type Argument1, Type? Argument2), Func<object>> s_registered = new();

    /// <summary>
    /// Registers how to create <paramref name="definition"/> of <paramref name="argument"/> without
    /// <c>MakeGenericType</c>.
    /// </summary>
    public static void Register(Type definition, Type argument, Func<object> create)
        => s_registered[(definition, argument, null)] = create;

    /// <summary>
    /// Registers how to create <paramref name="definition"/> of two type arguments without
    /// <c>MakeGenericType</c>.
    /// </summary>
    public static void Register(Type definition, Type argument1, Type argument2, Func<object> create)
        => s_registered[(definition, argument1, argument2)] = create;

    /// <summary>
    /// Creates <paramref name="definition"/> of <paramref name="argument"/> with its parameterless
    /// constructor.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
#if NET5_0_OR_GREATER
    // The generic types are made at run time, so the trimmer can't see that their constructors are needed.
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(TypeRule<>))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(NullableAdapter<>))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(CollectionRule<>))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(EnumMapper<>))]
#endif
#if NET7_0_OR_GREATER
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(StructNumberMapper<>))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(INumberBaseMapper<>))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(ParsableMapper<>))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(StructParsableMapper<>))]
#endif
    public static object Create(Type definition, Type argument)
        => s_registered.TryGetValue((definition, argument, null), out var create)
            ? create()
            : Activator.CreateInstance(definition.MakeGenericType(argument))!;

    /// <summary>
    /// Creates <paramref name="definition"/> of two type arguments with its parameterless constructor.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
#if NET5_0_OR_GREATER
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(DictionaryRule<,>))]
#endif
    public static object Create(Type definition, Type argument1, Type argument2)
        => s_registered.TryGetValue((definition, argument1, argument2), out var create)
            ? create()
            : Activator.CreateInstance(definition.MakeGenericType(argument1, argument2))!;
}
