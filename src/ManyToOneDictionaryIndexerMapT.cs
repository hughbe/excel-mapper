using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper;

/// <summary>
/// A map that reads one or more values from one or more cells and maps these values to the type of the
/// property or field. This is used to map IDictionary properties and fields.
/// </summary>
/// <typeparam name="TValue">The value type of the IDictionary property or field.</typeparam>
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
