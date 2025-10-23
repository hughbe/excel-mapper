using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a collection by calling the Add method to add items.
/// </summary>
/// <typeparam name="T">The type of the collection items.</typeparam>
public class AddEnumerableFactory<T> : IEnumerableFactory<T>
{
    public Type CollectionType { get; }
    private object? _items;
    private readonly MethodInfo _addMethod;

    /// <summary>
    /// Constructs a factory that creates collections of the given type.
    /// </summary>
    /// <param name="collectionType">The type of collection to create.</param>
    /// <exception cref="ArgumentException">Thrown when the collection type is invalid or unsupported.</exception>
    public AddEnumerableFactory(Type collectionType)
    {
        ArgumentNullException.ThrowIfNull(collectionType);
        if (collectionType.IsInterface)
        {
            throw new ArgumentException($"Interface list types cannot be created. Use {nameof(ListEnumerableFactory<T>)} instead.", nameof(collectionType));
        }
        if (collectionType.IsAbstract)
        {
            throw new ArgumentException("Abstract collection types cannot be created.", nameof(collectionType));
        }
        if (!collectionType.ImplementsInterface(typeof(IEnumerable)))
        {
            throw new ArgumentException($"Collection type {collectionType} must implement {nameof(IEnumerable)}.", nameof(collectionType));
        }
        if (collectionType.GetConstructor(Type.EmptyTypes) is null)
        {
            throw new ArgumentException($"Collection type {collectionType} must have a default constructor.", nameof(collectionType));
        }

        CollectionType = collectionType;
        _addMethod = collectionType.GetMethod("Add", [typeof(T)]) ?? throw new ArgumentException($"Type does not have an Add({typeof(T)}) method.", nameof(collectionType));
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = Activator.CreateInstance(CollectionType);
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        EnsureMapping();
        _addMethod.InvokeUnwrapped(_items, [item]);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        EnsureMapping();
        throw new NotSupportedException($"Set is not supported for {nameof(AddEnumerableFactory<T>)}.");
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
