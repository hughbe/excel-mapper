using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a ReadOnlyObservableCollection&lt;T&gt; by collecting items in an ObservableCollection and passing it to the constructor.
/// </summary>
/// <typeparam name="T">The type of the collection elements.</typeparam>
internal sealed class ReadOnlyObservableCollectionEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private readonly Func<object, object> _construct;
    private ObservableCollection<T?>? _items;

    /// <summary>
    /// Constructs a factory that creates collections of the given type.
    /// </summary>
    /// <param name="collectionType">The type of collection to create.</param>
    /// <exception cref="ArgumentException">Thrown when the collection type is invalid or unsupported.</exception>
    public ReadOnlyObservableCollectionEnumerableFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type collectionType)
    {
        var constructor = CollectionUtilities.GetConstructorTaking(collectionType, typeof(ObservableCollection<T>))!;
        _construct = CompiledAccessors.GetConstructor1(constructor);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _items = [];
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
        var factory = (ReadOnlyObservableCollectionEnumerableFactory<T>)MemberwiseClone();
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
