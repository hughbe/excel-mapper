using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a dictionary by passing a Dictionary&lt;TKey, TValue&gt; of items to a constructor that accepts IDictionary&lt;TKey, TValue&gt;.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
internal sealed class ConstructorDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue>, IBindingFactory where TKey : notnull
{
    private readonly Func<object, object> _construct;

    private Dictionary<TKey, TValue?>? _items;

    /// <summary>
    /// Constructs a factory that creates dictionaries of the given type.
    /// </summary>
    /// <param name="dictionaryType">The type of dictionary to create.</param>
    /// <exception cref="ArgumentException">Thrown when the dictionary type is invalid or unsupported.</exception>
    public ConstructorDictionaryFactory([DynamicallyAccessedMembers(TrimmingConstants.CollectionType)] Type dictionaryType)
    {
        var constructor = CollectionUtilities.GetConstructorTaking(dictionaryType, typeof(Dictionary<TKey, TValue>))!;
        _construct = CompiledAccessors.GetConstructor1(constructor);
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        _items = new Dictionary<TKey, TValue?>(count);
    }

    /// <inheritdoc/>
    public void Add(TKey key, TValue? value)
    {
        _items!.Add(key, value);
    }

    /// <inheritdoc/>
    public object End()
    {
        var items = _items!;
        _items = null;
        return _construct(items);
    }

    /// <inheritdoc/>
    /// <inheritdoc/>
    object IBindingFactory.CreateForBinding()
    {
        var factory = (ConstructorDictionaryFactory<TKey, TValue>)MemberwiseClone();
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
