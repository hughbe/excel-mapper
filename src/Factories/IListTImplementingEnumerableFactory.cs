using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a collection by instantiating a type that implements IList&lt;T&gt; and adding items.
/// </summary>
/// <typeparam name="T">The type of the collection elements.</typeparam>
public class IListTImplementingEnumerableFactory<T> : IEnumerableFactory<T>
{
    /// <summary>
    /// Gets the type of collection that this factory creates.
    /// </summary>
    public Type CollectionType { get; }
    private IList<T?>? _items;

    /// <summary>
    /// Constructs a factory that creates collections of the given type.
    /// </summary>
    /// <param name="collectionType">The type of collection to create.</param>
    /// <exception cref="ArgumentException">Thrown when the collection type is invalid or unsupported.</exception>
    public IListTImplementingEnumerableFactory(Type collectionType)
    {
        ThrowHelpers.ThrowIfNull(collectionType, nameof(collectionType));
        if (collectionType.IsInterface)
        {
            throw new ArgumentException($"Interface collection types cannot be created. Use {nameof(ListEnumerableFactory<T>)} instead.", nameof(collectionType));
        }
        if (collectionType.IsAbstract)
        {
            throw new ArgumentException("Abstract collection types cannot be created.", nameof(collectionType));
        }
        if (collectionType.IsArray)
        {
            throw new ArgumentException($"Array types cannot be created. Use {nameof(ArrayEnumerableFactory<T>)} instead.", nameof(collectionType));
        }
        if (!collectionType.ImplementsInterface(typeof(IList<T?>)))
        {
            throw new ArgumentException($"Collection type {collectionType} must implement {nameof(IList<T?>)}.", nameof(collectionType));
        }
        if (collectionType.GetConstructor(Type.EmptyTypes) is null)
        {
            throw new ArgumentException($"Collection type {collectionType} must have a default constructor.", nameof(collectionType));
        }

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

        _items = (IList<T?>)Activator.CreateInstance(CollectionType)!;
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
            return _items;
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
        if (_items is null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
