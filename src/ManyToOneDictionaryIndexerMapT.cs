using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to
/// a dictionary property or field by their keys.
/// </summary>
public class ManyToOneDictionaryIndexerMapT<TValue> : IDictionaryIndexerMap
{
    public ManyToOneDictionaryIndexerMapT(IDictionaryFactory<TValue> dictionaryFactory)
    {
        DictionaryFactory = dictionaryFactory ?? throw new ArgumentNullException(nameof(dictionaryFactory));
    }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    public IDictionaryFactory<TValue> DictionaryFactory { get; }

    /// <inheritdoc/>
    public Dictionary<string, IMap> Values { get; } = [];

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

        DictionaryFactory.Begin(Values.Count);
        try
        {
            foreach (var map in Values)
            {
                if (map.Value.TryGetValue(sheet, rowIndex, reader, member, out var elementValue))
                {
                    DictionaryFactory.Add(map.Key, (TValue)elementValue);
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
