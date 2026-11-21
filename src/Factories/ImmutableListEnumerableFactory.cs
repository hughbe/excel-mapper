using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs an ImmutableList&lt;T&gt; by using a builder and converting to immutable at the end.
/// </summary>
/// <typeparam name="T">The type of the list elements.</typeparam>
internal sealed class ImmutableListEnumerableFactory<T> : IEnumerableFactory<T>, IBindingFactory
{
    private ImmutableList<T?>.Builder? _builder;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _builder = ImmutableList.CreateBuilder<T?>();
    }

    /// <inheritdoc/>
    public void Add(T? item)
    {
        _builder!.Add(item);
    }

    /// <inheritdoc/>
    public void Set(int index, T? item)
    {
        // Grow the list if necessary.
        while (_builder!.Count <= index)
        {
            _builder.Add(default);
        }

        _builder[index] = item;
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
        var factory = (ImmutableListEnumerableFactory<T>)MemberwiseClone();
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
