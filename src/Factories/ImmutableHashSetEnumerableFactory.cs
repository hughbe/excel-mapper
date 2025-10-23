using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs an ImmutableHashSet&lt;T&gt; by using a builder and converting to immutable at the end.
/// </summary>
/// <typeparam name="T">The type of the set elements.</typeparam>
public class ImmutableHashSetEnumerableFactory<T> : IEnumerableFactory<T>
{
    private ImmutableHashSet<T?>.Builder? _builder;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_builder is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _builder = ImmutableHashSet.CreateBuilder<T?>();
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        EnsureMapping();
        _builder.Add(item);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        EnsureMapping();
        throw new NotSupportedException($"Set is not supported for {nameof(ImmutableHashSetEnumerableFactory<T>)}.");
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
