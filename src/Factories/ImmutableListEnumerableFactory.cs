using System;
using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class ImmutableListEnumerableFactory<T> : IEnumerableFactory<T>
{
    private ImmutableList<T?>.Builder? _builder;

    public void Begin(int count)
    {
        if (count < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(count), count, "Count cannot be negative.");
        }

        if (_builder is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _builder = ImmutableList.CreateBuilder<T?>();
    }

    public void Add(T? item)
    {
        EnsureMapping();
        _builder.Add(item);
    }

    public void Set(int index, T? item)
    {
        if (index < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(index), index, "Index cannot be negative.");
        }

        EnsureMapping();

        // Grow the list if necessary.
        while (_builder.Count <= index)
        {
            _builder.Add(default);
        }

        _builder[index] = item;
    }

    public object End()
    {
        EnsureMapping();

        var result = _builder.ToImmutable();
        Reset();
        return result;
    }

    public void Reset()
    {
        _builder = null;
    }

    [MemberNotNull(nameof(_builder))]
    private void EnsureMapping()
    {
        if (_builder is null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
