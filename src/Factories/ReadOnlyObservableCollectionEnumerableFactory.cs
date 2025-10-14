using System;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class ReadOnlyObservableCollectionEnumerableFactory<T> : IEnumerableFactory<T>
{
    public Type CollectionType { get; }
    private readonly ConstructorInfo _constructor;
    private ObservableCollection<T?>? _items;

    public ReadOnlyObservableCollectionEnumerableFactory(Type collectionType)
    {
        ArgumentNullException.ThrowIfNull(collectionType);
        if (collectionType.IsAbstract)
        {
            throw new ArgumentException("Abstract collection types cannot be created.", nameof(collectionType));
        }

        _constructor = collectionType.GetConstructor([typeof(ObservableCollection<T>)])
            ?? throw new ArgumentException($"Collection type {collectionType} does not have a constructor that takes {nameof(ObservableCollection<T>)}.", nameof(collectionType));
        CollectionType = collectionType;
    }

    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = [];
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
            return _constructor.Invoke([_items]);
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
        if (_items == null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
