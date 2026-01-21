using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a HashSet&lt;T&gt; by pre-allocating capacity and adding items.
/// </summary>
/// <typeparam name="T">The type of the set elements.</typeparam>
public class HashSetEnumerableFactory<T> : IEnumerableFactory<T>
{
    private HashSet<T?>? _items;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ThrowHelpers.ThrowIfNegative(count, nameof(count));

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = [];
#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP2_1_OR_GREATER
        _items.EnsureCapacity(count);
#endif
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
        throw new NotSupportedException($"Set is not supported for {nameof(HashSetEnumerableFactory<T>)}.");
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
        if (_items == null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
