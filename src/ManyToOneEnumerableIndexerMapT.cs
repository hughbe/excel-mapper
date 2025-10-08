using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper;

/// <summary>
/// A map that reads one or more values from one or more cells and maps these values to the type of the
/// property or field. This is used to map IEnumerable properties and fields.
/// </summary>
/// <typeparam name="TValue">The value type of the IEnumerable property or field.</typeparam>
public class ManyToOneEnumerableIndexerMapT<TValue> : IEnumerableIndexerMap
{
    public ManyToOneEnumerableIndexerMapT(IEnumerableFactory<TValue> enumerableFactory)
    {
        EnumerableFactory = enumerableFactory ?? throw new ArgumentNullException(nameof(enumerableFactory));
    }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    public IEnumerableFactory<TValue> EnumerableFactory { get; }

    /// <inheritdoc/>
    public Dictionary<int, IMap> Values { get; } = [];

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

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
