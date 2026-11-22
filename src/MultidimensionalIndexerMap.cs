using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to a
/// multidimensional array property or field by their indices.
/// </summary>
internal class MultidimensionalIndexerMap<TValue> : IMultidimensionalIndexerMap
{
    // Factories build one value at a time, so a map shared between importers (for example
    // a class map instance registered in several importers used concurrently) must not
    // build two values at once.
    private readonly object _buildLock = new();

    internal MultidimensionalIndexerMap(IMultidimensionalArrayFactory<TValue> arrayFactory)
    {
        ThrowHelpers.ThrowIfNull(arrayFactory, nameof(arrayFactory));
        ArrayFactory = arrayFactory;
    }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    internal IMultidimensionalArrayFactory<TValue> ArrayFactory { get; }

    /// <inheritdoc/>
    public Dictionary<int[], IMap> Values { get; } = [];

    private int[] GetLengths()
    {
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

    /// <inheritdoc/>
    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));

        var lengths = GetLengths();

        lock (_buildLock)
        {
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
}
