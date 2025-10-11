using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class DictionaryFactory<TValue> : IDictionaryFactory<TValue>
{
    private int _currentIndex = -1;
    private Dictionary<string, TValue?>? _items;

    public void Begin(int count)
    {
        if (_currentIndex != -1)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = new Dictionary<string, TValue?>(count);
        _currentIndex = 0;
    }

    public void Add(string key, TValue? value)
    {
        EnsureMapping();
        _items.Add(key, value);
        _currentIndex++;
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
        if (_items == null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }

        Debug.Assert(_currentIndex >= 0);
    }
}
