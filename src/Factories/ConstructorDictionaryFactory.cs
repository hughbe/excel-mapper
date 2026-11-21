using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a dictionary by passing a Dictionary&lt;TKey, TValue&gt; of items to a constructor that accepts IDictionary&lt;TKey, TValue&gt;.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
internal sealed class ConstructorDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue>, IBindingFactory where TKey : notnull
{
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
        var constructor = dictionaryType.GetConstructor([typeof(IDictionary<TKey, TValue>)])
            ?? throw new ArgumentException($"Dictionary type {dictionaryType} does not have a constructor that takes {nameof(IDictionary<TKey, TValue>)}.", nameof(dictionaryType));
        _construct = CompiledAccessors.GetConstructor1(constructor);
        _constructorCopies = CollectionUtilities.ConstructorCopies(dictionaryType);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
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
        _items!.Add(key, value);
    }

    /// <inheritdoc/>
    public object End()
    {
        try
        {
            return _construct(_items!);
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
}
