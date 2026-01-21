using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs an ImmutableArray&lt;T&gt; by using a builder and converting to immutable at the end.
/// </summary>
/// <typeparam name="T">The type of the array elements.</typeparam>
public class ImmutableArrayEnumerableFactory<T> : IEnumerableFactory<T>
{
    private ImmutableArray<T?>.Builder? _builder;
    private int _currentIndex = -1;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ThrowHelpers.ThrowIfNegative(count, nameof(count));

        if (_builder is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _builder = ImmutableArray.CreateBuilder<T?>(count);
        _builder.Count = count;
        _currentIndex = 0;
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        EnsureMapping();
        _builder[_currentIndex++] = item;
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        ThrowHelpers.ThrowIfNegative(index, nameof(index));
        EnsureMapping();
        ThrowHelpers.ThrowIfGreaterThanOrEqual(index, _builder.Count, nameof(index));
        _builder[index] = item;
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
