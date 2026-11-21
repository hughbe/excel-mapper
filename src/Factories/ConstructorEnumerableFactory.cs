using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a collection by passing a list of items to a constructor that accepts IList&lt;T&gt;, IEnumerable&lt;T&gt;, or ICollection.
/// </summary>
/// <typeparam name="T">The type of the collection elements.</typeparam>
public class ConstructorEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    /// <summary>
    /// Gets the type of collection that this factory creates.
    /// </summary>
    [DynamicallyAccessedMembers(TrimmingConstants.CollectionType)]
    public Type CollectionType { get; }
    private readonly Func<object, object> _construct;

    // Whether the constructor copies the collection passed to it, so the collection can be reused for the next row.
    private readonly bool _constructorCopies;

    private List<T?>? _items;
    private List<T?>? _buffer;

    /// <summary>
    /// Constructs a factory that creates collections of the given type.
    /// </summary>
    /// <param name="collectionType">The type of collection to create.</param>
    /// <exception cref="ArgumentException">Thrown when the collection type is invalid or unsupported.</exception>
    public ConstructorEnumerableFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type collectionType)
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

        var constructor = collectionType.GetConstructor([typeof(IList<T>)])
            ?? collectionType.GetConstructor([typeof(IEnumerable<T>)])
            ?? collectionType.GetConstructor([typeof(ICollection)])
            ?? throw new ArgumentException($"Collection type {collectionType} does not have a constructor that takes {nameof(IList<T>)}, {nameof(IEnumerable<T>)} or {nameof(ICollection)}.", nameof(collectionType));

        CollectionType = collectionType;
        _construct = CompiledAccessors.GetConstructor1(constructor);
        _constructorCopies = CollectionUtilities.ConstructorCopies(collectionType);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ThrowHelpers.ThrowIfNegative(count, nameof(count));

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

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
            return _construct(_items);
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
        var factory = (ConstructorEnumerableFactory<T>)MemberwiseClone();
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
        if (_items is not null && _constructorCopies)
        {
            // Clear the buffer so that it doesn't keep the row's values alive.
            _items.Clear();
            _buffer = _items;
        }

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
