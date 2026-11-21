using System.Collections.Frozen;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a FrozenSet&lt;T&gt; by collecting items in a list and then freezing them.
/// </summary>
/// <typeparam name="T">The type of the set elements.</typeparam>
internal sealed class FrozenSetEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private List<T?>? _items;
    private List<T?>? _buffer;

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
            _items = new List<T?>(count);
        }
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        _items!.Add(item);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
        // The type has no indexer, so members of it can't be mapped by index.
        => throw new NotSupportedException();

    /// <inheritdoc/>
    public object End()
    {
        try
        {
            return _items!.ToFrozenSet();
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
        var factory = (FrozenSetEnumerableFactory<T>)MemberwiseClone();
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
