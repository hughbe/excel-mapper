using System.Collections.Frozen;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a FrozenDictionary&lt;TKey, TValue&gt; by collecting key-value pairs in a dictionary and then freezing them.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
internal sealed class FrozenDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue>, IBindingFactory where TKey : notnull
{
    private Dictionary<TKey, TValue?>? _items;
    private Dictionary<TKey, TValue?>? _buffer;

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
            return _items!.ToFrozenDictionary();
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
        var factory = (FrozenDictionaryFactory<TKey, TValue>)MemberwiseClone();
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
        if (_items is not null)
        {
            // Clear the buffer so that it doesn't keep the row's values alive.
            _items.Clear();
            _buffer = _items;
        }

        _items = null;
    }
}
