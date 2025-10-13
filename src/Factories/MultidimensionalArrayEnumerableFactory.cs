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
        if (lengths == null)
        {
            throw new ArgumentNullException(nameof(lengths));
        }
        if (lengths.Length == 0)
        {
            throw new ArgumentException("Lengths cannot be empty.", nameof(lengths));
        }
        foreach (var length in lengths)
        {
            if (length < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(lengths), lengths, "Lengths cannot be negative.");
            }
        }

        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = Array.CreateInstance(typeof(T), lengths);
    }

    public void Set(int[] indices, T? item)
    {
        if (indices == null)
        {
            throw new ArgumentNullException(nameof(indices));
        }
        if (indices.Length == 0)
        {
            throw new ArgumentException("Indices cannot be empty.", nameof(indices));
        }
        foreach (var index in indices)
        {
            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(indices), indices, "Indices cannot be negative.");
            }
        }

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
