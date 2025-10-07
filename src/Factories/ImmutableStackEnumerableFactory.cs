using System.Collections.Generic;
using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class ImmutableStackEnumerableFactory<T> : IEnumerableFactory<T>
{
    private List<T?>? _items;

    public void Begin(int capacity)
    {
        if (_items is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _items = new List<T?>();
    }

    public void Add(T? item)
    {
        EnsureMapping();
        _items.Add(item);
    }

    public object End()
    {
        EnsureMapping();

        var result = _items;
        Reset();
        return ImmutableStack.CreateRange(result);
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
