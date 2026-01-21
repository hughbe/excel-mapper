using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a dictionary by passing a Dictionary&lt;TKey, TValue&gt; of items to a constructor that accepts IDictionary&lt;TKey, TValue&gt;.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
public class ConstructorDictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue> where TKey : notnull
{
    /// <summary>
    /// Gets the type of dictionary that this factory creates.
    /// </summary>
    public Type DictionaryType { get; }
    private readonly ConstructorInfo _constructor;

    private Dictionary<TKey, TValue?>? _items;

    /// <summary>
    /// Constructs a factory that creates dictionaries of the given type.
    /// </summary>
    /// <param name="dictionaryType">The type of dictionary to create.</param>
    /// <exception cref="ArgumentException">Thrown when the dictionary type is invalid or unsupported.</exception>
    public ConstructorDictionaryFactory(Type dictionaryType)
    {
        ThrowHelpers.ThrowIfNull(dictionaryType, nameof(dictionaryType));
        if (dictionaryType.IsInterface)
        {
            throw new ArgumentException("Interface dictionary types cannot be created. Use IDictionaryTImplementingFactory instead.", nameof(dictionaryType));
        }
        if (dictionaryType.IsAbstract)
        {
            throw new ArgumentException("Abstract dictionary types cannot be created.", nameof(dictionaryType));
        }

        _constructor = dictionaryType.GetConstructor([typeof(IDictionary<TKey, TValue>)])
            ?? throw new ArgumentException($"Dictionary type {dictionaryType} does not have a constructor that takes {nameof(IDictionary<TKey, TValue>)}.", nameof(dictionaryType));
        DictionaryType = dictionaryType;
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ThrowHelpers.ThrowIfNegative(count, nameof(count));

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = new Dictionary<TKey, TValue?>(count);
    }

    /// <inheritdoc/>
    public void Add(TKey key, TValue? value)
    {
        EnsureMapping();
        _items!.Add(key, value);
    }

    /// <inheritdoc/>
    public object End()
    {
        EnsureMapping();

        try
        {
            return _constructor.Invoke([_items]);
        }
        finally
        {
            Reset();
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        _items = null;
    }

    [MemberNotNull(nameof(_items))]
    private void EnsureMapping()
    {
        if (_items == null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
