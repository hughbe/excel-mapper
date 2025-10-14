using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to a
/// list property or field by their indices.
/// </summary>
public class ManyToOneEnumerableIndexerMapT<TValue> : IEnumerableIndexerMap
{
    public ManyToOneEnumerableIndexerMapT(IEnumerableFactory<TValue> enumerableFactory)
    {
        ArgumentNullException.ThrowIfNull(enumerableFactory);
        EnumerableFactory = enumerableFactory;
    }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    public IEnumerableFactory<TValue> EnumerableFactory { get; }

    /// <inheritdoc/>
    public Dictionary<int, IMap> Values { get; } = [];

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        EnumerableFactory.Begin(Math.Max(Values.Count > 0 ? Values.Max(x => x.Key + 1) : 0, Values.Count));
        try
        {
            foreach (var map in Values)
            {
                if (map.Value.TryGetValue(sheet, rowIndex, reader, member, out var elementValue))
                {
                    EnumerableFactory.Set(map.Key, (TValue)elementValue);
                }
            }

            value = EnumerableFactory.End();
            return true;
        }
        finally
        {
            EnumerableFactory.Reset();
        }
    }
}
