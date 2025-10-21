using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Factories;

public class IListTImplementingEnumerableFactory<T> : IEnumerableFactory<T>
{
    public Type CollectionType { get; }
    private IList<T?>? _items;

    public IListTImplementingEnumerableFactory(Type collectionType)
    {
        ArgumentNullException.ThrowIfNull(collectionType);
        if (collectionType.IsInterface)
        {
            throw new ArgumentException($"Interface collection types cannot be created. Use {nameof(ListEnumerableFactory<T>)} instead.", nameof(collectionType));
        }
        if (collectionType.IsAbstract)
        {
            throw new ArgumentException("Abstract collection types cannot be created.", nameof(collectionType));
        }
        if (collectionType.IsArray)
        {
            throw new ArgumentException($"Array types cannot be created. Use {nameof(ArrayEnumerableFactory<T>)} instead.", nameof(collectionType));
        }
        if (!collectionType.ImplementsInterface(typeof(IList<T?>)))
        {
            throw new ArgumentException($"Collection type {collectionType} must implement {nameof(IList<T?>)}.", nameof(collectionType));
        }
        if (collectionType.GetConstructor(Type.EmptyTypes) is null)
        {
            throw new ArgumentException($"Collection type {collectionType} must have a default constructor.", nameof(collectionType));
        }

        CollectionType = collectionType;
    }

    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = (IList<T?>)Activator.CreateInstance(CollectionType)!;
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
            _items.Add(default);
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
