using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a collection by instantiating a type that implements IList and adding items.
/// </summary>
/// <typeparam name="T">The type of the collection elements.</typeparam>
internal sealed class IListImplementingEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private readonly Func<object> _create;
    private IList? _items;

    /// <summary>
    /// Constructs a factory that creates lists of the given type.
    /// </summary>
    /// <param name="listType">The type of list to create.</param>
    /// <exception cref="ArgumentException">Thrown when the list type is invalid or unsupported.</exception>
    public IListImplementingEnumerableFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type listType)
    {
        _create = CompiledAccessors.GetParameterlessConstructor(listType);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _items = (IList)_create();
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
            _items.Add(default(T));
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
        var factory = (IListImplementingEnumerableFactory<T>)MemberwiseClone();
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
