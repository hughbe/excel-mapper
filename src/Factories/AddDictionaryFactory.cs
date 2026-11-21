using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a dictionary by calling the Add method to add key-value pairs.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
public class AddDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue>, IBindingFactory where TKey : notnull
{
    /// <summary>
    /// Gets the type of dictionary that this factory creates.
    /// </summary>
    [DynamicallyAccessedMembers(TrimmingConstants.CollectionType)]
    public Type DictionaryType { get; }
    private object? _items;
    private readonly Func<object> _create;
    private readonly Action<object, TKey, TValue?> _add;

    /// <summary>
    /// Constructs a factory that creates dictionaries of the given type.
    /// </summary>
    /// <param name="dictionaryType">The type of dictionary to create.</param>
    /// <exception cref="ArgumentException">Thrown when the dictionary type is invalid or unsupported.</exception>
    public AddDictionaryFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type dictionaryType)
    {
        ThrowHelpers.ThrowIfNull(dictionaryType, nameof(dictionaryType));
        if (dictionaryType.IsInterface)
        {
            throw new ArgumentException("Interface dictionary types cannot be created. Use DictionaryEnumerableFactory instead.", nameof(dictionaryType));
        }
        if (dictionaryType.IsAbstract)
        {
            throw new ArgumentException("Abstract dictionary types cannot be created.", nameof(dictionaryType));
        }
        if (!dictionaryType.ImplementsInterface(typeof(IEnumerable)))
        {
            throw new ArgumentException($"Dictionary type {dictionaryType} must implement {nameof(IEnumerable)}.", nameof(dictionaryType));
        }
        if (dictionaryType.GetConstructor(Type.EmptyTypes) is null)
        {
            throw new ArgumentException($"Dictionary type {dictionaryType} must have a default constructor.", nameof(dictionaryType));
        }

        DictionaryType = dictionaryType;
        var addMethod = dictionaryType.GetMethod("Add", [typeof(TKey), typeof(TValue)]) ?? throw new ArgumentException($"Type does not have an Add({typeof(TKey)}, {typeof(TValue)}) method.", nameof(dictionaryType));
        _create = CompiledAccessors.GetParameterlessConstructor(dictionaryType);
        _add = CompiledAccessors.GetAction<TKey, TValue?>(addMethod);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ThrowHelpers.ThrowIfNegative(count, nameof(count));

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = _create();
    }

    /// <inheritdoc/>
    public void Add(TKey key, TValue? value)
    {
        EnsureMapping();
        _add(_items, key, value);
    }

    /// <inheritdoc/>
    public object End()
    {
        EnsureMapping();

        try
        {
            return _items;
        }
        finally
        {
            Reset();
        }
    }

    /// <inheritdoc/>
    /// <inheritdoc/>
    object IBindingFactory.CreateForBinding()
    {
        var factory = (AddDictionaryFactory<TKey, TValue>)MemberwiseClone();
        factory.ResetForBinding();
        return factory;
    }

    // A binding's copy starts with nothing being built and no buffer of its own. Only drop the
    // references: the objects may be in use by the map's own factory on another thread.
    private void ResetForBinding()
    {
        _items = null;
    }

    public void Reset()
    {
        _items = null;
    }

    [MemberNotNull(nameof(_items))]
    private void EnsureMapping()
    {
        if (_items == null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
