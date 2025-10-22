using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Factories;

public class DictionaryFactory<TKey, TValue> : IDictionaryFactory<TKey, TValue> where TKey : notnull
{
    private Dictionary<TKey, TValue?>? _items;

    public void Begin(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        if (_items is not null)
        {
            throw new ExcelMappingException($"Cannot begin mapping until {nameof(End)}() was called.");
        }

        _items = new Dictionary<TKey, TValue?>(count);
    }

    public void Add(TKey key, TValue? value)
    {
        ArgumentNullException.ThrowIfNull(key);
        EnsureMapping();
        _items.Add(key, value);
    }

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
