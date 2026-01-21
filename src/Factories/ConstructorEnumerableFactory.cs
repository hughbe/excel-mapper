using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a collection by passing a list of items to a constructor that accepts IList&lt;T&gt;, IEnumerable&lt;T&gt;, or ICollection.
/// </summary>
/// <typeparam name="T">The type of the collection elements.</typeparam>
public class ConstructorEnumerableFactory<T> : IEnumerableFactory<T>
{
    /// <summary>
    /// Gets the type of collection that this factory creates.
    /// </summary>
    public Type CollectionType { get; }
    private readonly ConstructorInfo _constructor;

    private List<T?>? _items;

    /// <summary>
    /// Constructs a factory that creates collections of the given type.
    /// </summary>
    /// <param name="collectionType">The type of collection to create.</param>
    /// <exception cref="ArgumentException">Thrown when the collection type is invalid or unsupported.</exception>
    public ConstructorEnumerableFactory(Type collectionType)
    {
        ThrowHelpers.ThrowIfNull(collectionType, nameof(collectionType));
        if (collectionType.IsInterface)
        {
            throw new ArgumentException("Interface collection types cannot be created. Use ListEnumerableFactory instead.", nameof(collectionType));
        }
        if (collectionType.IsAbstract)
        {
            throw new ArgumentException("Abstract collection types cannot be created.", nameof(collectionType));
        }

        _constructor = collectionType.GetConstructor([typeof(IList<T>)])
            ?? collectionType.GetConstructor([typeof(IEnumerable<T>)])
            ?? collectionType.GetConstructor([typeof(ICollection)])
            ?? throw new ArgumentException($"Collection type {collectionType} does not have a constructor that takes {nameof(IList<T>)}, {nameof(IEnumerable<T>)} or {nameof(ICollection)}.", nameof(collectionType));

        CollectionType = collectionType;
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ThrowHelpers.ThrowIfNegative(count, nameof(count));

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = new List<T?>(count);
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
        ThrowHelpers.ThrowIfNegative(index, nameof(index));
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
