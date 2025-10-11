using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class ConstructorDictionaryFactory<T> : IDictionaryFactory<T>
{
    public Type DictionaryType { get; }
    private readonly ConstructorInfo _constructor;

    private Dictionary<string, T?>? _items;

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

        _constructor = dictionaryType.GetConstructor([typeof(IDictionary<string, T>)])
            ?? throw new ArgumentException($"Dictionary type {dictionaryType} does not have a constructor that takes IDictionary<{typeof(T)}>.", nameof(dictionaryType));
        DictionaryType = dictionaryType;
    }

    public void Begin(int count)
    {
        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = new Dictionary<string, T?>(count);
    }

    public void Add(string key, T? value)
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
