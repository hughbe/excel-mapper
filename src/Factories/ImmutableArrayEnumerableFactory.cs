using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs an ImmutableArray&lt;T&gt; by using a builder and converting to immutable at the end.
/// </summary>
/// <typeparam name="T">The type of the array elements.</typeparam>
internal sealed class ImmutableArrayEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private ImmutableArray<T?>.Builder? _builder;
    private int _currentIndex = -1;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _builder = ImmutableArray.CreateBuilder<T?>(count);
        _builder.Count = count;
        _currentIndex = 0;
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        _builder![_currentIndex++] = item;
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        ThrowHelpers.ThrowIfGreaterThanOrEqual(index, _builder!.Count, nameof(index));
        _builder[index] = item;
    }

    /// <inheritdoc/>
    public object End()
    {
        try
        {
            // The builder was created at the final size, so its array becomes the result without a copy.
            return _builder!.Capacity == _builder.Count ? _builder.MoveToImmutable() : _builder.ToImmutable();
        }
        finally
        {
            Reset();
        }
    }

    /// <inheritdoc/>
    /// <inheritdoc/>
    object IBindingFactory.CreateForBinding()
    {
        var factory = (ImmutableArrayEnumerableFactory<T>)MemberwiseClone();
        factory.ResetForBinding();
        return factory;
    }

    // A binding's copy starts with nothing being built and no buffer of its own. Only drop the
    // references: the objects may be in use by the map's own factory on another thread.
    private void ResetForBinding()
    {
        _builder = null;
        _currentIndex = -1;
    }

    public void Reset()
    {
        _builder = null;
    }
}
