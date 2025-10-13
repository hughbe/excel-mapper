using System;
using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class ImmutableSortedSetEnumerableFactory<T> : IEnumerableFactory<T>
{
    private ImmutableSortedSet<T?>.Builder? _builder;

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

        _builder = ImmutableSortedSet.CreateBuilder<T?>();
    }

    public void Add(T? item)
    {
        EnsureMapping();
        _builder.Add(item);
    }

    public void Set(int index, T? item)
    {
        EnsureMapping();
        throw new NotSupportedException("Set is not supported for ImmutableSortedSetEnumerableFactory.");
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
