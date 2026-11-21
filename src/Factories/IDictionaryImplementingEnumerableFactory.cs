using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a dictionary by instantiating a type that implements IDictionary and adding key-value pairs.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
internal sealed class IDictionaryImplementingFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue>, IBindingFactory where TKey : notnull
{
    private readonly Func<object> _create;
    private IDictionary? _items;

    /// <summary>
    /// Constructs a factory that creates dictionaries of the given type.
    /// </summary>
    /// <param name="dictionaryType">The type of dictionary to create.</param>
    /// <exception cref="ArgumentException">Thrown when the dictionary type is invalid or unsupported.</exception>
    public IDictionaryImplementingFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type dictionaryType)
    {
        _create = CompiledAccessors.GetParameterlessConstructor(dictionaryType);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _items = (IDictionary)_create();
    }

    /// <inheritdoc/>
    public void Add(TKey key, TValue? value)
    {
        _items!.Add(key, value);
    }

    /// <inheritdoc/>
    public object End()
    {
        try
        {
            return _items!;
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
        var factory = (IDictionaryImplementingFactory<TKey, TValue>)MemberwiseClone();
        factory.ResetForBinding();
        return factory;
    }

    // A binding's copy starts with nothing being built and no buffer of its own. Only drop the
    // references: the objects may be in use by the map's own factory on another thread.
    private void ResetForBinding()
    {
        _items = null;
    }

    public void Reset()
    {
        _items = null;
    }
}
