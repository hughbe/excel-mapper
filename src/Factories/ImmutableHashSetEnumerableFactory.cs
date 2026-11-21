using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs an ImmutableHashSet&lt;T&gt; by using a builder and converting to immutable at the end.
/// </summary>
/// <typeparam name="T">The type of the set elements.</typeparam>
internal sealed class ImmutableHashSetEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private ImmutableHashSet<T?>.Builder? _builder;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _builder = ImmutableHashSet.CreateBuilder<T?>();
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        _builder!.Add(item);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        throw new NotSupportedException($"Set is not supported for {nameof(ImmutableHashSetEnumerableFactory<T>)}.");
    }

    /// <inheritdoc/>
    public object End()
    {
        try
        {
            return _builder!.ToImmutable();
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
        var factory = (ImmutableHashSetEnumerableFactory<T>)MemberwiseClone();
        factory.ResetForBinding();
        return factory;
    }

    // A binding's copy starts with nothing being built and no buffer of its own. Only drop the
    // references: the objects may be in use by the map's own factory on another thread.
    private void ResetForBinding()
    {
        _builder = null;
    }

    public void Reset()
    {
        _builder = null;
    }
}
