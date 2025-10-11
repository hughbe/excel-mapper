using System;
using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Factories;

public class IDictionaryImplementingFactory<T> : IDictionaryFactory<T>
{
    public Type DictionaryType { get; }
    private IDictionary? _items;

    public IDictionaryImplementingFactory(Type dictionaryType)
    {
        if (dictionaryType is null)
        {
            throw new ArgumentNullException(nameof(dictionaryType));
        }
        if (dictionaryType.IsInterface)
        {
            throw new ArgumentException("Interface dictionary types cannot be created. Use DictionaryEnumerableFactory instead.", nameof(dictionaryType));
        }
        if (dictionaryType.IsAbstract)
        {
            throw new ArgumentException("Abstract dictionary types cannot be created.", nameof(dictionaryType));
        }
        if (!dictionaryType.ImplementsInterface(typeof(IDictionary)))
        {
            throw new ArgumentException($"Dictionary type {dictionaryType} must implement IDictionary.", nameof(dictionaryType));
        }

        DictionaryType = dictionaryType;
    }

    public void Begin(int count)
    {
        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = (IDictionary)Activator.CreateInstance(DictionaryType)!;
    }

    public void Add(string key, T? value)
    {
        EnsureMapping();
        _items.Add(key, value);
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
