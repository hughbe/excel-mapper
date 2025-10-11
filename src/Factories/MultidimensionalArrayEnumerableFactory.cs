using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class MultidimensionalArrayFactory<T> : IMultidimensionalArrayFactory<T>
{
    private Array? _items;

    public void Begin(int[] lengths)
    {
        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = Array.CreateInstance(typeof(T), lengths);
    }

    public void Set(int[] indices, T? item)
    {
        EnsureMapping();
        _items.SetValue(item, indices);
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
        if (_items is null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
