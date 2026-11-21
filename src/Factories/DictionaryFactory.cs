using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a Dictionary&lt;TKey, TValue&gt; by pre-allocating capacity and adding key-value pairs.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
internal sealed class DictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue>, IBindingFactory where TKey : notnull
{
    private Dictionary<TKey, TValue?>? _items;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _items = new Dictionary<TKey, TValue?>(count);
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
            return _items!;
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
        var factory = (DictionaryFactory<TKey, TValue>)MemberwiseClone();
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
}
