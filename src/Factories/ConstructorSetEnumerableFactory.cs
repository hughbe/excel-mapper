using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a set by passing a HashSet&lt;T&gt; of items to a constructor that accepts ISet&lt;T&gt;.
/// </summary>
/// <typeparam name="T">The type of the set elements.</typeparam>
public class ConstructorSetEnumerableFactory<T> : IEnumerableFactory<T>
{
    public Type SetType { get; }
    private readonly ConstructorInfo _constructor;

    private HashSet<T?>? _items;

    /// <summary>
    /// Constructs a factory that creates sets of the given type.
    /// </summary>
    /// <param name="setType">The type of set to create.</param>
    /// <exception cref="ArgumentException">Thrown when the set type is invalid or unsupported.</exception>
    public ConstructorSetEnumerableFactory(Type setType)
    {
        ArgumentNullException.ThrowIfNull(setType);
        if (setType.IsInterface)
        {
            throw new ArgumentException("Interface set types cannot be created. Use HashSetEnumerableFactory instead.", nameof(setType));
        }
        if (setType.IsAbstract)
        {
            throw new ArgumentException("Abstract set types cannot be created.", nameof(setType));
        }

        _constructor = setType.GetConstructor(new[] { typeof(ISet<T>) })
            ?? throw new ArgumentException($"Set type {setType} does not have a constructor that takes {nameof(ISet<T?>)}.", nameof(setType));

        SetType = setType;
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = new HashSet<T?>(count);
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
        EnsureMapping();
        throw new NotSupportedException($"Set is not supported for {nameof(ConstructorSetEnumerableFactory<T>)}.");
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
