using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to a
/// list property or field by their indices.
/// </summary>
public class ManyToOneEnumerableIndexerMapT<TValue> : IEnumerableIndexerMap
{
    // Factories build one value at a time, so a map shared between importers (for example
    // a class map instance registered in several importers used concurrently) must not
    // build two values at once.
    private readonly object _buildLock = new();

    public ManyToOneEnumerableIndexerMapT(IEnumerableFactory<TValue> enumerableFactory)
    {
        ThrowHelpers.ThrowIfNull(enumerableFactory, nameof(enumerableFactory));
        EnumerableFactory = enumerableFactory;
    }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    public IEnumerableFactory<TValue> EnumerableFactory { get; }

    /// <inheritdoc/>
    public Dictionary<int, IMap> Values { get; } = [];

    private int GetLength()
    {
        // The collection must be large enough for the largest index.
        var length = Values.Count;
        foreach (var index in Values.Keys)
        {
            length = Math.Max(length, index + 1);
        }

        return length;
    }

    /// <inheritdoc/>
    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));

        lock (_buildLock)
        {
            EnumerableFactory.Begin(GetLength());
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
}
