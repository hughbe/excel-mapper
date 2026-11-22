using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using ExcelMapper.Mappers;

namespace ExcelMapper.AutoMapping;

/// <summary>
/// Creates the generic types auto-mapping needs for types found at run time, such as the
/// <see cref="TypeRule{T}"/> of a member's type or the <see cref="EnumMapper{TEnum}"/> of an enum.
/// They're made with <c>MakeGenericType</c>, which native AOT can't do for types it didn't compile,
/// unless an <see cref="ExcelMappingContext"/> registered them (see <see cref="ExcelTypeRegistrar"/>):
/// a registration creates the type with <c>new</c>, which native AOT compiles.
/// </summary>
/// <remarks>
/// Each type created here for a type found at run time needs an <see cref="ExcelTypeRegistrar"/>
/// method that registers it, which the source generator's TypeWalker calls: see "Auto-mapping and
/// the source generator" in CLAUDE.md.
/// </remarks>
internal static class GenericTypes
{
    private static readonly ConcurrentDictionary<(Type Definition, Type Argument1, Type? Argument2), Func<object>> s_registered = new();

    /// <summary>
    /// Gets or sets whether a mapping context throws when it maps a type it didn't register, under
    /// the JIT as with native AOT. The tests turn it on, so that a generated context whose generator
    /// and auto-mapping disagree fails them rather than a published app (see "Auto-mapping and the
    /// source generator" in CLAUDE.md). It's internal, for the tests only, which set it by reflection
    /// (tests/RequireRegisteredTypes.cs): don't rename it without changing them.
    /// </summary>
    internal static bool RequireRegisteredTypes { get; set; }

    // The context creating a class map on this thread, if any, and the generic types it registered:
    // only its class maps are checked, against its own registrations (other contexts' registrations
    // are global, but a context that relied on them would fail in an app without them).
    [ThreadStatic]
    private static Type? t_context;

    [ThreadStatic]
    private static HashSet<(Type, Type, Type?)>? t_registered;

    // What each instance created here is, to check it's registered each time it's used.
    private static readonly ConditionalWeakTable<object, Tuple<Type, Type, Type?>> s_created = new();

    /// <summary>
    /// Marks the class maps created until the scope ends as those of <paramref name="contextType"/>,
    /// checked by <see cref="Require(Type, Type, Type?)"/>.
    /// </summary>
    public static ContextScope EnterContext(Type contextType, HashSet<(Type, Type, Type?)> registeredTypes)
    {
        var scope = new ContextScope(t_context, t_registered);
        t_context = contextType;
        t_registered = registeredTypes;
        return scope;
    }

    /// <summary>
    /// Ends <see cref="EnterContext"/>.
    /// </summary>
    public readonly struct ContextScope(Type? previousContext, HashSet<(Type, Type, Type?)>? previousRegistered) : IDisposable
    {
        public void Dispose()
        {
            t_context = previousContext;
            t_registered = previousRegistered;
        }
    }

    /// <summary>
    /// With <see cref="RequireRegisteredTypes"/> on, while a mapping context creates a class
    /// map, throws if <paramref name="definition"/> of the arguments isn't registered. Called each
    /// time a generic type made for a type found at run time is used, not only when it's created,
    /// so one created earlier (and cached) is checked too.
    /// </summary>
    /// <exception cref="InvalidOperationException">The type isn't registered.</exception>
    public static void Require(Type definition, Type argument1, Type? argument2 = null)
    {
        if (RequireRegisteredTypes && t_context is { } context && t_registered is { } registered && !registered.Contains((definition, argument1, argument2)))
        {
            throw new InvalidOperationException(
                $"The mapping context {context} doesn't register {Describe(definition, argument1, argument2)}, which its class maps need. " +
                $"Register it with ExcelTypeRegistrar in RegisterTypes, or let the source generator write the context.");
        }
    }

    /// <summary>
    /// Checks <paramref name="instance"/>, if it was created here, as <see cref="Require(Type, Type, Type?)"/>.
    /// </summary>
    public static void Require(object? instance)
    {
        if (RequireRegisteredTypes && instance is not null && s_created.TryGetValue(instance, out var key))
        {
            Require(key.Item1, key.Item2, key.Item3);
        }
    }

    // What registers a generic type, in the words of the ExcelTypeRegistrar method that does.
    private static string Describe(Type definition, Type argument1, Type? argument2)
        => definition == typeof(TypeRule<>) ? $"the type {argument1} (Type, Nullable, Enum, Collection, Array, Dictionary or Parsable)" :
           definition == typeof(NullableAdapter<>) ? $"the nullable type {argument1}? (Nullable)" :
           definition == typeof(CollectionRule<>) ? $"collections of {argument1} (Collection or Array)" :
           definition == typeof(DictionaryRule<,>) ? $"dictionaries of {argument1} to {argument2} (Dictionary)" :
           definition == typeof(EnumMapper<>) ? $"the enum {argument1} (Enum)" :
           $"the type {argument1} ({definition.Name})";

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
        => Remember(
            s_registered.TryGetValue((definition, argument, null), out var create)
                ? create()
                : Activator.CreateInstance(definition.MakeGenericType(argument))!,
            definition, argument, null);

    /// <summary>
    /// Creates <paramref name="definition"/> of two type arguments with its parameterless constructor.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
#if NET5_0_OR_GREATER
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors, typeof(DictionaryRule<,>))]
#endif
    public static object Create(Type definition, Type argument1, Type argument2)
        => Remember(
            s_registered.TryGetValue((definition, argument1, argument2), out var create)
                ? create()
                : Activator.CreateInstance(definition.MakeGenericType(argument1, argument2))!,
            definition, argument1, argument2);

    // Remembers what the instance is, for Require. A registered factory may return an instance
    // that already exists (such as TypeRule<T>.Instance), which keeps what it was first created as.
    private static object Remember(object instance, Type definition, Type argument1, Type? argument2)
    {
        if (RequireRegisteredTypes)
        {
            s_created.GetValue(instance, _ => Tuple.Create(definition, argument1, argument2));
        }

        Require(definition, argument1, argument2);
        return instance;
    }
}
