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
        ArgumentNullException.ThrowIfNull(lengths);
        if (lengths.Length == 0)
        {
            throw new ArgumentException("Lengths cannot be empty.", nameof(lengths));
        }
        foreach (var length in lengths)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(length, nameof(lengths));
        }

        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = Array.CreateInstance(typeof(T), lengths);
    }

    public void Set(int[] indices, T? item)
    {
        ArgumentNullException.ThrowIfNull(indices);
        if (indices.Length == 0)
        {
            throw new ArgumentException("Indices cannot be empty.", nameof(indices));
        }
        foreach (var index in indices)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index, nameof(indices));
        }

        EnsureMapping();
        _items.SetValue(item, indices);
    }

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
