using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.AutoMapping;
using ExcelMapper.Mappers;

namespace ExcelMapper;

/// <summary>
/// Registers the types an <see cref="ExcelMappingContext"/> maps, so that auto-mapping creates the
/// generic code it needs for them directly, which native AOT compiles, rather than with
/// <c>MakeGenericType</c>, which native AOT can't do. Used by
/// <see cref="ExcelMappingContext.RegisterTypes"/>: register every type auto-mapping meets.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
// Registering a type only records how to create its generic types. They're created by
// GenericTypes.Create, during auto-mapping, whose callers carry the warnings.
[UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "The registered factories are only called by GenericTypes.Create, which requires unreferenced code.")]
[UnconditionalSuppressMessage("AOT", "IL3050", Justification = "The registered factories are only called by GenericTypes.Create, which requires dynamic code.")]
public sealed class ExcelTypeRegistrar
{
    internal ExcelTypeRegistrar()
    {
    }

    internal HashSet<Type> RowTypes { get; } = [];

    // The generic types registered, which the context's class maps may use (see GenericTypes.Require).
    internal HashSet<(Type Definition, Type Argument1, Type? Argument2)> RegisteredTypes { get; } = [];

    private void Register(Type definition, Type argument, Func<object> create)
    {
        GenericTypes.Register(definition, argument, create);
        RegisteredTypes.Add((definition, argument, null));
    }

    private void Register(Type definition, Type argument1, Type argument2, Func<object> create)
    {
        GenericTypes.Register(definition, argument1, argument2, create);
        RegisteredTypes.Add((definition, argument1, argument2));
    }

    /// <summary>
    /// Registers a type rows are mapped to, whose class map
    /// <see cref="ExcelMappingContext.GetClassMap{T}"/> gives. Register it with
    /// <see cref="Collection{TCollection, TElement}"/>, <see cref="Array{TElement}"/> or
    /// <see cref="Dictionary{TDictionary, TKey, TValue}"/> too if it's a collection or dictionary.
    /// </summary>
    /// <typeparam name="T">The type.</typeparam>
    public void Row<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
    {
        Type<T>();
        RowTypes.Add(typeof(T));
    }

    /// <summary>
    /// Registers the type of a member, or of an element or value, that has a built-in conversion
    /// or is mapped through its own members.
    /// </summary>
    /// <typeparam name="T">The type.</typeparam>
    public void Type<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
        => Register(typeof(TypeRule<>), typeof(T), static () => TypeRule<T>.Instance);

    /// <summary>
    /// Registers the nullable version of a value type, such as <c>int?</c> or the nullable version of
    /// an enum.
    /// </summary>
    /// <typeparam name="T">The value type.</typeparam>
    public void Nullable<T>() where T : struct
    {
        Register(typeof(TypeRule<>), typeof(T?), static () => TypeRule<T?>.Instance);
        Register(typeof(NullableAdapter<>), typeof(T), static () => new NullableAdapter<T>());
    }

    /// <summary>
    /// Registers an enum, mapped from its names, numbers and descriptions. Register its nullable
    /// version with <see cref="Nullable{T}"/>.
    /// </summary>
    /// <typeparam name="TEnum">The enum.</typeparam>
    public void Enum<TEnum>() where TEnum : struct, Enum
    {
        Register(typeof(TypeRule<>), typeof(TEnum), static () => TypeRule<TEnum>.Instance);
        Register(typeof(EnumMapper<>), typeof(TEnum), static () => new EnumMapper<TEnum>());
    }

    /// <summary>
    /// Registers a collection type other than an array, such as <c>List&lt;int&gt;</c>, and its
    /// element type's collections. Register the element type too if it isn't a built-in type.
    /// </summary>
    /// <typeparam name="TCollection">The collection type.</typeparam>
    /// <typeparam name="TElement">The element type.</typeparam>
    public void Collection<[DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] TCollection, TElement>()
    {
        Register(typeof(TypeRule<>), typeof(TCollection), static () => TypeRule<TCollection>.Instance);
        Register(typeof(CollectionRule<>), typeof(TElement), static () => new CollectionRule<TElement>());
    }

    /// <summary>
    /// Registers an array type, such as <c>int[]</c>. Register the element type too if it isn't a
    /// built-in type.
    /// </summary>
    /// <typeparam name="TElement">The element type.</typeparam>
    public void Array<TElement>()
    {
        Register(typeof(TypeRule<>), typeof(TElement[]), static () => TypeRule<TElement[]>.Instance);
        Register(typeof(CollectionRule<>), typeof(TElement), static () => new CollectionRule<TElement>());
    }

    /// <summary>
    /// Registers a dictionary type, such as <c>Dictionary&lt;string, int&gt;</c>. Register the value
    /// type too if it isn't a built-in type.
    /// </summary>
    /// <typeparam name="TDictionary">The dictionary type.</typeparam>
    /// <typeparam name="TKey">The key type.</typeparam>
    /// <typeparam name="TValue">The value type.</typeparam>
    public void Dictionary<[DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] TDictionary, TKey, TValue>() where TKey : notnull
    {
        Register(typeof(TypeRule<>), typeof(TDictionary), static () => TypeRule<TDictionary>.Instance);
        Register(typeof(DictionaryRule<,>), typeof(TKey), typeof(TValue), static () => new DictionaryRule<TKey, TValue>());
    }

    /// <summary>
    /// Registers how to set a property or field of a <typeparamref name="TOwner"/>, as ordinary
    /// code, so that it's set without compiling an expression (which native AOT can't) or
    /// reflection (which boxes values and allocates).
    /// </summary>
    /// <typeparam name="TOwner">The type whose member is set.</typeparam>
    /// <typeparam name="TValue">The member's type.</typeparam>
    /// <param name="member">The name of the property or field.</param>
    /// <param name="setter">Sets the member.</param>
    /// <exception cref="ArgumentNullException"><paramref name="member"/> or <paramref name="setter"/> is null.</exception>
    public void Setter<TOwner, TValue>(string member, ExcelMemberSetter<TOwner, TValue> setter)
    {
        ThrowHelpers.ThrowIfNull(member, nameof(member));
        ThrowHelpers.ThrowIfNull(setter, nameof(setter));
        RegisteredAccessors.RegisterSetter(member, setter);
    }

    /// <summary>
    /// Registers how to call a public constructor of a <typeparamref name="TOwner"/>, as ordinary
    /// code, so that types created with constructor parameters (such as records) are created
    /// without compiling an expression or reflection.
    /// </summary>
    /// <typeparam name="TOwner">The type created.</typeparam>
    /// <param name="parameterTypes">The types of the constructor's parameters, which identify it.</param>
    /// <param name="constructor">Calls the constructor.</param>
    /// <exception cref="ArgumentNullException"><paramref name="parameterTypes"/> or <paramref name="constructor"/> is null.</exception>
    public void Constructor<TOwner>(Type[] parameterTypes, ExcelConstructor<TOwner> constructor)
    {
        ThrowHelpers.ThrowIfNull(parameterTypes, nameof(parameterTypes));
        ThrowHelpers.ThrowIfNull(constructor, nameof(constructor));
        RegisteredAccessors.RegisterConstructor(parameterTypes, constructor);
    }

#if NET7_0_OR_GREATER
    /// <summary>
    /// Registers a value type parsed with <see cref="IParsable{TSelf}"/> that isn't a number, mapped
    /// from the cell's text. Register a reference type with <see cref="ParsableClass{T}"/>.
    /// </summary>
    /// <typeparam name="T">The type.</typeparam>
    public void Parsable<T>() where T : struct, IParsable<T>
    {
        Register(typeof(TypeRule<>), typeof(T), static () => TypeRule<T>.Instance);
        Register(typeof(StructParsableMapper<>), typeof(T), static () => new StructParsableMapper<T>());
    }

    /// <summary>
    /// Registers a reference type parsed with <see cref="IParsable{TSelf}"/> that isn't a number,
    /// mapped from the cell's text.
    /// </summary>
    /// <typeparam name="T">The type.</typeparam>
    public void ParsableClass<T>() where T : class, IParsable<T>
    {
        Register(typeof(TypeRule<>), typeof(T), static () => TypeRule<T>.Instance);
        Register(typeof(ParsableMapper<>), typeof(T), static () => new ParsableMapper<T>());
    }
#endif
}
