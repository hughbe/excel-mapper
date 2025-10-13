using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Factories;

public class AddDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue> where TKey : notnull
{
    public Type DictionaryType { get;  }
    private object? _items;
    private readonly MethodInfo _addMethod;

    public AddDictionaryFactory(Type dictionaryType)
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
        if (!dictionaryType.ImplementsInterface(typeof(IEnumerable)))
        {
            throw new ArgumentException($"Dictionary type {dictionaryType} must implement IEnumerable.", nameof(dictionaryType));
        }

        DictionaryType = dictionaryType;
        _addMethod = dictionaryType.GetMethod("Add", [typeof(TKey), typeof(TValue)]) ?? throw new ArgumentException($"Type does not have an Add({typeof(TKey)}, {typeof(TValue)}) method.", nameof(dictionaryType));
    }

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

        _items = Activator.CreateInstance(DictionaryType);
    }

    public void Add(TKey key, TValue? value)
    {
        EnsureMapping();
        _addMethod.InvokeUnwrapped(_items, [key, value]);
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
