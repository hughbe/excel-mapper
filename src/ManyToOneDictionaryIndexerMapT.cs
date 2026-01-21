using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to
/// a dictionary property or field by their keys.
/// </summary>
public class ManyToOneDictionaryIndexerMapT<TKey, TValue> : IDictionaryIndexerMap where TKey : notnull
{
    public ManyToOneDictionaryIndexerMapT(IDictionaryFactory<TKey, TValue> dictionaryFactory)
    {
        ThrowHelpers.ThrowIfNull(dictionaryFactory, nameof(dictionaryFactory));
        DictionaryFactory = dictionaryFactory;
    }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    public IDictionaryFactory<TKey, TValue> DictionaryFactory { get; }

    /// <inheritdoc/>
    public Dictionary<object, IMap> Values { get; } = [];

    /// <inheritdoc/>
    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));

        DictionaryFactory.Begin(Values.Count);
        try
        {
            foreach (var map in Values)
            {
                if (map.Value.TryGetValue(sheet, rowIndex, reader, member, out var elementValue))
                {
                    DictionaryFactory.Add((TKey)map.Key, (TValue?)elementValue);
                }
            }

            value = DictionaryFactory.End();
            return true;
        }
        finally
        {
            DictionaryFactory.Reset();
        }
    }
}
