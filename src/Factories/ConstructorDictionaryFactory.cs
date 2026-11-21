using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a dictionary by passing a Dictionary&lt;TKey, TValue&gt; of items to a constructor that accepts IDictionary&lt;TKey, TValue&gt;.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
public class ConstructorDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue>, IBindingFactory where TKey : notnull
{
    /// <summary>
    /// Gets the type of dictionary that this factory creates.
    /// </summary>
    [DynamicallyAccessedMembers(TrimmingConstants.CollectionType)]
    public Type DictionaryType { get; }
    private readonly Func<object, object> _construct;

    // Whether the constructor copies the collection passed to it, so the collection can be reused for the next row.
    private readonly bool _constructorCopies;

    private Dictionary<TKey, TValue?>? _items;
    private Dictionary<TKey, TValue?>? _buffer;

    /// <summary>
    /// Constructs a factory that creates dictionaries of the given type.
    /// </summary>
    /// <param name="dictionaryType">The type of dictionary to create.</param>
    /// <exception cref="ArgumentException">Thrown when the dictionary type is invalid or unsupported.</exception>
    public ConstructorDictionaryFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type dictionaryType)
    {
        ThrowHelpers.ThrowIfNull(dictionaryType, nameof(dictionaryType));
        if (dictionaryType.IsInterface)
        {
            throw new ArgumentException("Interface dictionary types cannot be created. Use IDictionaryTImplementingFactory instead.", nameof(dictionaryType));
        }
        if (dictionaryType.IsAbstract)
        {
            throw new ArgumentException("Abstract dictionary types cannot be created.", nameof(dictionaryType));
        }

        var constructor = dictionaryType.GetConstructor([typeof(IDictionary<TKey, TValue>)])
            ?? throw new ArgumentException($"Dictionary type {dictionaryType} does not have a constructor that takes {nameof(IDictionary<TKey, TValue>)}.", nameof(dictionaryType));
        DictionaryType = dictionaryType;
        _construct = CompiledAccessors.GetConstructor1(constructor);
        _constructorCopies = CollectionUtilities.ConstructorCopies(dictionaryType);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ThrowHelpers.ThrowIfNegative(count, nameof(count));

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        if (_buffer is not null)
        {
            // The buffer from the previous row: the result is a copy of it, so it can be reused.
            _items = _buffer;
            _buffer = null;
        }
        else
        {
            _items = new Dictionary<TKey, TValue?>(count);
        }
    }

    /// <inheritdoc/>
    public void Add(TKey key, TValue? value)
    {
        EnsureMapping();
        _items!.Add(key, value);
    }

    /// <inheritdoc/>
    public object End()
    {
        EnsureMapping();

        try
        {
            return _construct(_items);
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
        var factory = (ConstructorDictionaryFactory<TKey, TValue>)MemberwiseClone();
        factory.ResetForBinding();
        return factory;
    }

    // A binding's copy starts with nothing being built and no buffer of its own. Only drop the
    // references: the objects may be in use by the map's own factory on another thread.
    private void ResetForBinding()
    {
        _items = null;
        _buffer = null;
    }

    public void Reset()
    {
        if (_items is not null && _constructorCopies)
        {
            // Clear the buffer so that it doesn't keep the row's values alive.
            _items.Clear();
            _buffer = _items;
        }

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
