using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Factories;

public class ImmutableDictionaryFactory<TValue> : IDictionaryFactory<TValue>
{
    private ImmutableDictionary<string, TValue?>.Builder? _builder;

    public void Begin(int count)
    {
        if (_builder is not null)
        {
            throw new ExcelMappingException("Cannot begin mapping until End() was called.");
        }

        _builder = ImmutableDictionary.CreateBuilder<string, TValue?>();
    }

    public void Add(string key, TValue? value)
    {
        EnsureMapping();
        _builder.Add(key, value);
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
