using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a collection by instantiating a type that implements IList and adding items.
/// </summary>
/// <typeparam name="T">The type of the collection elements.</typeparam>
public class IListImplementingEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    /// <summary>
    /// Gets the type of list that this factory creates.
    /// </summary>
    [DynamicallyAccessedMembers(TrimmingConstants.CollectionType)]
    public Type ListType { get; }

    private readonly Func<object> _create;
    private IList? _items;

    /// <summary>
    /// Constructs a factory that creates lists of the given type.
    /// </summary>
    /// <param name="listType">The type of list to create.</param>
    /// <exception cref="ArgumentException">Thrown when the list type is invalid or unsupported.</exception>
    public IListImplementingEnumerableFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type listType)
    {
        ThrowHelpers.ThrowIfNull(listType, nameof(listType));
        if (listType.IsInterface)
        {
            throw new ArgumentException($"Interface collection types cannot be created. Use {nameof(ListEnumerableFactory<T>)} instead.", nameof(listType));
        }
        if (listType.IsAbstract)
        {
            throw new ArgumentException("Abstract list types cannot be created.", nameof(listType));
        }
        if (listType.IsArray)
        {
            throw new ArgumentException($"Array types cannot be created. Use {nameof(ArrayEnumerableFactory<T>)} instead.", nameof(listType));
        }
        if (!listType.ImplementsInterface(typeof(IList)))
        {
            throw new ArgumentException($"List type {listType} must implement {nameof(IList)}.", nameof(listType));
        }
        if (listType.GetConstructor(Type.EmptyTypes) is null)
        {
            throw new ArgumentException($"List type {listType} must have a default constructor.", nameof(listType));
        }

        ListType = listType;
        _create = CompiledAccessors.GetParameterlessConstructor(listType);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ThrowHelpers.ThrowIfNegative(count, nameof(count));

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = (IList)_create();
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
            _items.Add(default(T));
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

    [MemberNotNull(nameof(_items))]
    private void EnsureMapping()
    {
        if (_items is null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
