using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Utilities;

namespace ExcelMapper.Factories;

public class IDictionaryTImplementingFactory<TKey, TValue> : Abstractions.IDictionaryFactory<TKey, TValue> where TKey : notnull
{
    public Type DictionaryType { get; }
    private IDictionary<TKey, TValue?>? _items;

    public IDictionaryTImplementingFactory(Type dictionaryType)
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
        if (!dictionaryType.ImplementsInterface(typeof(IDictionary<TKey, TValue?>)))
        {
            throw new ArgumentException($"Dictionary type {dictionaryType} must implement IDictionary<{typeof(TKey)}, {typeof(TValue)}>.", nameof(dictionaryType));
        }

        DictionaryType = dictionaryType;
    }

    public void Begin(int count)
    {
        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = (IDictionary<TKey, TValue?>)Activator.CreateInstance(DictionaryType)!;
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
