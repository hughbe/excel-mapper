using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Factories;

public class ISetTImplementingEnumerableFactory<T> : IEnumerableFactory<T>
{
    public Type SetType { get; }
    private ISet<T?>? _items;

    public ISetTImplementingEnumerableFactory(Type setType)
    {
        ArgumentNullException.ThrowIfNull(setType);
        if (setType.IsInterface)
        {
            throw new ArgumentException("Interface set types cannot be created. Use HashSetEnumerableFactory instead.", nameof(setType));
        }
        if (setType.IsAbstract)
        {
            throw new ArgumentException("Abstract set types cannot be created.", nameof(setType));
        }
        if (!setType.ImplementsInterface(typeof(ISet<T?>)))
        {
            throw new ArgumentException($"Set type {setType} must implement {nameof(ISet<T?>)}.", nameof(setType));
        }
        if (setType.GetConstructor(Type.EmptyTypes) is null)
        {
            throw new ArgumentException($"Set type {setType} must have a default constructor.", nameof(setType));
        }

        SetType = setType;
    }

    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = (ISet<T?>)Activator.CreateInstance(SetType)!;
    }

    public void Add(T? item)
    {
        EnsureMapping();
        _items.Add(item);
    }

    public void Set(int index, T? item)
    {
        EnsureMapping();
        throw new NotSupportedException($"Set is not supported for {nameof(HashSetEnumerableFactory<T>)}.");
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
        if (_items == null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
