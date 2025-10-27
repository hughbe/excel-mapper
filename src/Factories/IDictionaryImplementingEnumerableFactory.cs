using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

/// <summary>
/// Constructs a dictionary by instantiating a type that implements IDictionary and adding key-value pairs.
/// </summary>
/// <typeparam name="TKey">The type of the dictionary keys.</typeparam>
/// <typeparam name="TValue">The type of the dictionary values.</typeparam>
public class IDictionaryImplementingFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue> where TKey : notnull
{
    /// <summary>
    /// Gets the type of dictionary that this factory creates.
    /// </summary>
    public Type DictionaryType { get; }
    private IDictionary? _items;

    /// <summary>
    /// Constructs a factory that creates dictionaries of the given type.
    /// </summary>
    /// <param name="dictionaryType">The type of dictionary to create.</param>
    /// <exception cref="ArgumentException">Thrown when the dictionary type is invalid or unsupported.</exception>
    public IDictionaryImplementingFactory(Type dictionaryType)
    {
        ArgumentNullException.ThrowIfNull(dictionaryType);
        if (dictionaryType.IsInterface)
        {
            throw new ArgumentException($"Interface dictionary types cannot be created. Use {nameof(DictionaryFactory<TKey, TValue>)} instead.", nameof(dictionaryType));
        }
        if (dictionaryType.IsAbstract)
        {
            throw new ArgumentException("Abstract dictionary types cannot be created.", nameof(dictionaryType));
        }
        if (!dictionaryType.ImplementsInterface(typeof(IDictionary)))
        {
            throw new ArgumentException($"Dictionary type {dictionaryType} must implement {nameof(IDictionary)}.", nameof(dictionaryType));
        }
        if (dictionaryType.GetConstructor(Type.EmptyTypes) is null)
        {
            throw new ArgumentException($"Dictionary type {dictionaryType} must have a default constructor.", nameof(dictionaryType));
        }

        DictionaryType = dictionaryType;
    }

    /// <inheritdoc/>
    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = (IDictionary)Activator.CreateInstance(DictionaryType)!;
    }

    /// <inheritdoc/>
    public void Add(TKey key, TValue? value)
    {
        ArgumentNullException.ThrowIfNull(key);
        EnsureMapping();
        _items.Add(key, value);
    }

    /// <inheritdoc/>
    public object End()
    {
        EnsureMapping();

        try
        {
            return _items;
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
        if (_items is null)
        {
            throw new ExcelMappingException("Has not started mapping.");
        }
    }
}
