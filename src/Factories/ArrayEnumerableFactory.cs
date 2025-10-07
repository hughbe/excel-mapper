using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class ArrayEnumerableFactory<T> : IEnumerableFactory<T>
{
    private int _currentIndex = -1;
    private T?[]? _items;

    public void Begin(int capacity)
    {
        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = new T[capacity];
        _currentIndex = 0;
    }

    public void Add(T? item)
    {
        EnsureMapping();
        _items[_currentIndex++] = item;
    }

    public object End()
    {
        EnsureMapping();

        var result = _items;
        Reset();
        return result;
    }

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
