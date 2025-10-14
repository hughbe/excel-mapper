using System;
using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Factories;

public class AddEnumerableFactory<T> : IEnumerableFactory<T>
{
    public Type CollectionType { get;  }
    private object? _items;
    private readonly MethodInfo _addMethod;

    public AddEnumerableFactory(Type collectionType)
    {
        ArgumentNullException.ThrowIfNull(collectionType);
        if (collectionType.IsInterface)
        {
            throw new ArgumentException("Interface list types cannot be created. Use ListEnumerableFactory instead.", nameof(collectionType));
        }
        if (collectionType.IsAbstract)
        {
            throw new ArgumentException("Abstract collection types cannot be created.", nameof(collectionType));
        }
        if (!collectionType.ImplementsInterface(typeof(IEnumerable)))
        {
            throw new ArgumentException($"Collection type {collectionType} must implement IEnumerable.", nameof(collectionType));
        }
        if (collectionType.GetConstructor(Type.EmptyTypes) is null)
        {
            throw new ArgumentException($"Collection type {collectionType} must have a default constructor.", nameof(collectionType));
        }

        CollectionType = collectionType;
        _addMethod = collectionType.GetMethod("Add", [typeof(T)]) ?? throw new ArgumentException($"Type does not have an Add({typeof(T)}) method.", nameof(collectionType));
    }

    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = Activator.CreateInstance(CollectionType);
    }

    public void Add(T? item)
    {
        EnsureMapping();
        _addMethod.InvokeUnwrapped(_items, [item]);
    }

    public void Set(int index, T? item)
    {
        EnsureMapping();
        throw new NotSupportedException("Set is not supported for AddEnumerableFactory.");
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
