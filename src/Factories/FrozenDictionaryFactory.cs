using System;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class FrozenDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue> where TKey : notnull
{
    private Dictionary<TKey, TValue?>? _items;

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

        _items = new Dictionary<TKey, TValue?>(count);
    }

    public void Add(TKey key, TValue? value)
    {
        EnsureMapping();
        _items.Add(key, value);
    }

    public object End()
    {
        EnsureMapping();

        var result = _items;
        Reset();
        return result.ToFrozenDictionary();
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
