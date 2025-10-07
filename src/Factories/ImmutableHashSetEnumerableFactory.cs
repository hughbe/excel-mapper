using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class ImmutableHashSetEnumerableFactory<T> : IEnumerableFactory<T>
{
    private ImmutableHashSet<T?>.Builder? _builder;

    public void Begin(int capacity)
    {
        if (_builder is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _builder = ImmutableHashSet.CreateBuilder<T?>();
    }

    public void Add(T? item)
    {
        EnsureMapping();
        _builder.Add(item);
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
