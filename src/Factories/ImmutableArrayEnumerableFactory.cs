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
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_builder is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
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
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        EnsureMapping();
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _builder.Count);
        _builder[index] = item;
    }

    public object End()
    {
        EnsureMapping();

        try
        {
            return _builder.ToImmutable();
        }
        finally
        {
            Reset();
        }
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
