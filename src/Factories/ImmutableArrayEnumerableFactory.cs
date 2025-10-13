using System;
using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class ImmutableArrayEnumerableFactory<T> : IEnumerableFactory<T>
{
    private ImmutableArray<T?>.Builder? _builder;
    private int _currentIndex = -1;

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

        _builder = ImmutableArray.CreateBuilder<T?>(count);
        _builder.Count = count;
        _currentIndex = 0;
    }

    public void Add(T? item)
    {
        EnsureMapping();
        _builder[_currentIndex++] = item;
    }

    public void Set(int index, T? item)
    {
        EnsureMapping();
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
