using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class ConstructorDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue> where TKey : notnull
{
    public Type DictionaryType { get; }
    private readonly ConstructorInfo _constructor;

    private Dictionary<TKey, TValue?>? _items;

    public ConstructorDictionaryFactory(Type dictionaryType)
    {
        if (dictionaryType == null)
        {
            throw new ArgumentNullException(nameof(dictionaryType));
        }
        if (dictionaryType.IsInterface)
        {
            throw new ArgumentException("Interface dictionary types cannot be created. Use IDictionaryTImplementingFactory instead.", nameof(dictionaryType));
        }
        if (dictionaryType.IsAbstract)
        {
            throw new ArgumentException("Abstract dictionary types cannot be created.", nameof(dictionaryType));
        }

        _constructor = dictionaryType.GetConstructor([typeof(IDictionary<TKey, TValue>)])
            ?? throw new ArgumentException($"Dictionary type {dictionaryType} does not have a constructor that takes IDictionary<{typeof(TKey)}, {typeof(TValue)}>.", nameof(dictionaryType));
        DictionaryType = dictionaryType;
    }

    public void Begin(int count)
    {
        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = new Dictionary<TKey, TValue?>(count);
    }

    public void Add(TKey key, TValue? value)
    {
        EnsureMapping();
        _items!.Add(key, value);
    }

    public object End()
    {
        EnsureMapping();

        var result = _constructor.Invoke([_items]);
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
