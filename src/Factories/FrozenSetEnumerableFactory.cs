using System.Collections.Frozen;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a FrozenSet&lt;T&gt; by collecting items in a list and then freezing them.
/// </summary>
/// <typeparam name="T">The type of the set elements.</typeparam>
public class FrozenSetEnumerableFactory<T> : IEnumerableFactory<T>
{
    private List<T?>? _items;

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
            return _items.ToFrozenSet();
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
