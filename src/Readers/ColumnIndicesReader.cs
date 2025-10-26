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
    public IReadOnlyList<int> ColumnIndices { get; }

    /// <summary>
    /// Constructs a reader that reads the values of cells at the given column indices.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices for each column to read.</param>
    public ColumnIndicesReader(params IReadOnlyList<int> columnIndices)
    {
        ColumnUtilities.ValidateColumnIndices(columnIndices);
        ColumnIndices = columnIndices;
    }

    /// <inheritdoc/>
    public bool TryGetValues(IExcelDataReader reader, bool preserveFormatting, [NotNullWhen(true)] out IEnumerable<ReadCellResult>? result)
    {
        result = ColumnIndices.Select(columnIndex => new ReadCellResult(columnIndex, reader, preserveFormatting));
        return true;
    }
}
