using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs an ImmutableSortedDictionary&lt;TKey, TValue&gt; by using a builder and converting to immutable at the end.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
public class ImmutableSortedDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue> where TKey : notnull
{
    private ImmutableSortedDictionary<TKey, TValue?>.Builder? _builder;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_builder is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _builder = ImmutableSortedDictionary.CreateBuilder<TKey, TValue?>();
    }

    /// <inheritdoc/>
    public void Add(TKey key, TValue? value)
    {
        ArgumentNullException.ThrowIfNull(key);
        EnsureMapping();
        _builder.Add(key, value);
    }

    /// <inheritdoc/>
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

    /// <inheritdoc/>
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
