using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs an array by pre-allocating storage and adding items sequentially.
/// </summary>
/// <typeparam name="T">The type of the array elements.</typeparam>
public class ArrayEnumerableFactory<T> : IEnumerableFactory<T>
{
    private int _currentIndex = -1;
    private T?[]? _items;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = new T[count];
        _currentIndex = 0;
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        EnsureMapping();
        _items[_currentIndex++] = item;
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        EnsureMapping();
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _items.Length);

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
        _currentIndex = -1;
    }

    [MemberNotNull(nameof(_items))]
    private void EnsureMapping()
    {
        if (_items is null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }

        Debug.Assert(_currentIndex >= 0);
    }
}
