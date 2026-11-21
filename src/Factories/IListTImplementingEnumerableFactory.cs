using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a collection by instantiating a type that implements IList&lt;T&gt; and adding items.
/// </summary>
/// <typeparam name="T">The type of the collection elements.</typeparam>
internal sealed class IListTImplementingEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private readonly Func<object> _create;
    private IList<T?>? _items;

    /// <summary>
    /// Constructs a factory that creates collections of the given type.
    /// </summary>
    /// <param name="collectionType">The type of collection to create.</param>
    /// <exception cref="ArgumentException">Thrown when the collection type is invalid or unsupported.</exception>
    public IListTImplementingEnumerableFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type collectionType)
    {
        _create = CompiledAccessors.GetParameterlessConstructor(collectionType);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _items = (IList<T?>)_create();
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        _items!.Add(item);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        // Grow the list if necessary.
        while (_items!.Count <= index)
        {
            _items.Add(default);
        }

        _items[index] = item;
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
        var factory = (IListTImplementingEnumerableFactory<T>)MemberwiseClone();
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
