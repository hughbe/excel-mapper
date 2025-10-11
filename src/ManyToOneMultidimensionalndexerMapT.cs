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
/// multidimensional array property or field by their indices.
/// </summary>
public class ManyToOneMultidimensionalIndexerMapT<TValue> : IMultidimensionalIndexerMap
{
    public ManyToOneMultidimensionalIndexerMapT(IMultidimensionalArrayFactory<TValue> arrayFactory)
    {
        ArrayFactory = arrayFactory ?? throw new ArgumentNullException(nameof(arrayFactory));
    }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    public IMultidimensionalArrayFactory<TValue> ArrayFactory { get; }

    /// <inheritdoc/>
    public Dictionary<int[], IMap> Values { get; } = [];

    private int[] GetLengths()
    {
        if (Values.Count == 0)
        {
            return [];
        }

        // Get the rank (number of dimensions) from any key
        int rank = Values.Keys.First().Length;

        // Calculate the length needed for each dimension (max index + 1)
        int[] lengths = new int[rank];
        foreach (var indices in Values.Keys)
        {
            for (int i = 0; i < rank; i++)
            {
                lengths[i] = Math.Max(lengths[i], indices[i] + 1);
            }
        }

        return lengths;
    }

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

        var lengths = GetLengths();
        if (lengths.Length == 0)
        {
            value = null;
            return false;
        }

        ArrayFactory.Begin(lengths);
        try
        {
            foreach (var map in Values)
            {
                if (map.Value.TryGetValue(sheet, rowIndex, reader, member, out var elementValue))
                {
                    ArrayFactory.Set(map.Key, (TValue)elementValue);
                }
            }

            value = ArrayFactory.End();
            return true;
        }
        finally
        {
            ArrayFactory.Reset();
        }
    }
}
