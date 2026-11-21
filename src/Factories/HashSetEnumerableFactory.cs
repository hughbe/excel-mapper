using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a HashSet&lt;T&gt; by pre-allocating capacity and adding items.
/// </summary>
/// <typeparam name="T">The type of the set elements.</typeparam>
internal sealed class HashSetEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private HashSet<T?>? _items;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _items = [];
#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP2_1_OR_GREATER
        _items.EnsureCapacity(count);
#endif
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        _items!.Add(item);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        throw new NotSupportedException($"Set is not supported for {nameof(HashSetEnumerableFactory<T>)}.");
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
        var factory = (HashSetEnumerableFactory<T>)MemberwiseClone();
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
