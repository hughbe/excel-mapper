using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs an ImmutableSortedDictionary&lt;TKey, TValue&gt; by using a builder and converting to immutable at the end.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
internal sealed class ImmutableSortedDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue>, IBindingFactory where TKey : notnull
{
    private ImmutableSortedDictionary<TKey, TValue?>.Builder? _builder;

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _builder = ImmutableSortedDictionary.CreateBuilder<TKey, TValue?>();
    }

    /// <inheritdoc/>
    public void Add(TKey key, TValue? value)
    {
        _builder!.Add(key, value);
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
        var factory = (ImmutableSortedDictionaryFactory<TKey, TValue>)MemberwiseClone();
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
