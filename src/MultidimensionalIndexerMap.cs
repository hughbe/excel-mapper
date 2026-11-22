using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to a
/// multidimensional array property or field by their indices.
/// </summary>
internal class MultidimensionalIndexerMap<TValue> : IParentMap
{
    [RequiresDynamicCode("Creates multidimensional arrays of TValue whose rank is only known at run time, which native AOT can't guarantee.")]
    internal MultidimensionalIndexerMap()
    {
    }

    public Dictionary<int[], IMap> Values { get; } = [];

    IEnumerable<KeyValuePair<object?, IMap>> IParentMap.Children
    {
        get
        {
            foreach (var value in Values)
            {
                yield return new(value.Key, value.Value);
            }
        }
    }

    IMap? IParentMap.GetChild(object key) => FindKey((int[])key) is { } existing ? Values[existing] : null;

    void IParentMap.SetChild(object key, IMap map)
    {
        ThrowHelpers.ThrowIfFrozen(_frozen);
        var indices = (int[])key;
        Values[FindKey(indices) ?? indices] = map;
    }

    private bool _frozen;

    /// <inheritdoc />
    void IMap.Freeze()
    {
        if (_frozen)
        {
            return;
        }

        _frozen = true;
        foreach (var map in Values.Values)
        {
            map.Freeze();
        }
    }

    // Arrays don't implement equality, so the key is found by its indices.
    private int[]? FindKey(int[] indices)
    {
        foreach (var key in Values.Keys)
        {
            if (key.SequenceEqual(indices))
            {
                return key;
            }
        }

        return null;
    }

    string? IParentMap.GetChildPath(string? path, object? key) => $"{path}[{string.Join(", ", (int[])key!)}]";

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
    public bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? value)
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
