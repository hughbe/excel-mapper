using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a ReadOnlyObservableCollection&lt;T&gt; by collecting items in an ObservableCollection and passing it to the constructor.
/// </summary>
/// <typeparam name="T">The type of the collection elements.</typeparam>
public class ReadOnlyObservableCollectionEnumerableFactory<T> : IEnumerableFactory<T>
{
    public Type CollectionType { get; }
    private readonly ConstructorInfo _constructor;
    private ObservableCollection<T?>? _items;

    /// <summary>
    /// Constructs a factory that creates collections of the given type.
    /// </summary>
    /// <param name="collectionType">The type of collection to create.</param>
    /// <exception cref="ArgumentException">Thrown when the collection type is invalid or unsupported.</exception>
    public ReadOnlyObservableCollectionEnumerableFactory(Type collectionType)
    {
        ArgumentNullException.ThrowIfNull(collectionType);
        if (collectionType.IsAbstract)
        {
            throw new ArgumentException("Abstract collection types cannot be created.", nameof(collectionType));
        }

        _constructor = collectionType.GetConstructor([typeof(ObservableCollection<T>)])
            ?? throw new ArgumentException($"Collection type {collectionType} does not have a constructor that takes {nameof(ObservableCollection<T>)}.", nameof(collectionType));
        CollectionType = collectionType;
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = [];
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        EnsureMapping();
        _items.Add(item);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        EnsureMapping();

        // Grow the list if necessary.
        while (_items.Count <= index)
        {
            _items.Add(default);
        }

        _items[index] = item;
    }

    /// <inheritdoc/>
    public object End()
    {
        EnsureMapping();

        try
        {
            return _constructor.Invoke([_items]);
        }
        finally
        {
            Reset();
        }
    }

    /// <inheritdoc/>
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
