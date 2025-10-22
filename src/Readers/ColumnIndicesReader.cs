using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a cell.
/// </summary>
public class ColumnIndicesReader : ICellsReader
{
    /// <summary>
    /// Gets the zero-based indices for each column to read.
    /// </summary>
    public IList<int> ColumnIndices { get; }

    public ColumnIndicesReader(IList<int> columnIndices)
    {
        ColumnUtilities.ValidateColumnIndices(columnIndices, nameof(columnIndices));
        ColumnIndices = columnIndices;
    }

    public bool TryGetValues(IExcelDataReader reader, bool preserveFormatting, [NotNullWhen(true)] out IEnumerable<ReadCellResult>? result)
    {
        result = ColumnIndices.Select(columnIndex => new ReadCellResult(columnIndex, reader, preserveFormatting));
        return true;
    }
}
