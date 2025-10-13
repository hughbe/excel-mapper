using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class HashSetEnumerableFactory<T> : IEnumerableFactory<T>
{
    private HashSet<T?>? _items;

    public void Begin(int count)
    {
        if (count < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(count), count, "Count cannot be negative.");
        }

        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = [];
        _items.EnsureCapacity(count);
    }

    public void Add(T? item)
    {
        EnsureMapping();
        _items.Add(item);
    }

    public void Set(int index, T? item)
    {
        EnsureMapping();
        throw new NotSupportedException("Set is not supported for HashSetEnumerableFactory.");
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
