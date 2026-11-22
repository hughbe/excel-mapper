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
    [RequiresDynamicCode("Creates multidimensional arrays of TValue whose rank is only known at run time, which native AOT can't guarantee.")]
    internal MultidimensionalIndexerMap()
    {
    }

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
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "The constructor requires dynamic code.")]
    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));

        var array = Array.CreateInstance(typeof(TValue), GetLengths());
        foreach (var map in Values)
        {
            if (map.Value.TryGetValue(sheet, rowIndex, reader, member, errors, out var elementValue))
            {
                array.SetValue((TValue)elementValue, map.Key);
            }
        }

        value = array;
        return true;
    }
}
