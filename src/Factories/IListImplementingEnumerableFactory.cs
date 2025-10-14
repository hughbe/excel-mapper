using System;
using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Factories;

public class IListImplementingEnumerableFactory<T> : IEnumerableFactory<T>
{
    public Type ListType { get; }
    private IList? _items;

    public IListImplementingEnumerableFactory(Type listType)
    {
        ArgumentNullException.ThrowIfNull(listType);
        if (listType.IsInterface)
        {
            throw new ArgumentException("Interface collection types cannot be created. Use ListEnumerableFactory instead.", nameof(listType));
        }
        if (listType.IsAbstract)
        {
            throw new ArgumentException("Abstract list types cannot be created.", nameof(listType));
        }
        if (listType.IsArray)
        {
            throw new ArgumentException("Array types cannot be created. Use ArrayEnumerableFactory instead.", nameof(listType));
        }
        if (!listType.ImplementsInterface(typeof(IList)))
        {
            throw new ArgumentException($"List type {listType} must implement IList.", nameof(listType));
        }

        ListType = listType;
    }

    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = (IList)Activator.CreateInstance(ListType)!;
    }

    public void Add(T? item)
    {
        EnsureMapping();
        _items.Add(item);
    }

    public void Set(int index, T? item)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        EnsureMapping();

        // Grow the list if necessary.
        while (_items.Count <= index)
        {
            _items.Add(default(T));
        }

        _items[index] = item;
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
