using System.Diagnostics.CodeAnalysis;
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

    private IExcelDataReader? _reader;
    private bool _preserveFormatting;
    private int _currentIndex;

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
    public bool Start(IExcelDataReader reader, bool preserveFormatting, out int count)
    {
        _reader = reader;
        _preserveFormatting = preserveFormatting;
        _currentIndex = -1;
        count = ColumnIndices.Count;
        return true;
    }

    /// <inheritdoc/>
    public bool TryGetNext([NotNullWhen(true)] out ReadCellResult result)
    {
        if (_currentIndex < ColumnIndices.Count - 1)
        {
            _currentIndex++;
            result = new ReadCellResult(ColumnIndices[_currentIndex], _reader!, _preserveFormatting);
            return true;
        }

        result = default;
        return false;
    }

    /// <inheritdoc/>
    public void Reset()
    {
        _currentIndex = -1;
        _reader = null;
    }
}
