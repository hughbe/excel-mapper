using System.Diagnostics.CodeAnalysis;
using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// A cells reader that combines multiple cell readers.
/// </summary>
public class CompositeCellsReader : ICellsReader
{
    /// <summary>
    /// The cell readers.
    /// </summary>
    public IReadOnlyList<ICellReader> Readers { get; }

    /// <summary>
    /// Initializes a new instance of <see cref="CompositeCellsReader"/>.
    /// </summary>
    /// <param name="readers">The cell readers.</param>
    /// <exception cref="ArgumentException">Thrown when the readers list is empty or contains null values.</exception>
    public CompositeCellsReader(params IReadOnlyList<ICellReader> readers)
    {
        ArgumentNullException.ThrowIfNull(readers);
        if (readers.Count == 0)
        {
            throw new ArgumentException("At least one reader must be provided.", nameof(readers));
        }
        foreach (var reader in readers)
        {
            if (reader == null)
            {
                throw new ArgumentException("Readers cannot contain null values.", nameof(readers));
            }
        }

        Readers = readers;
    }

    /// <inheritdoc/>
    public bool TryGetValues(IExcelDataReader reader, bool preserveFormatting, [NotNullWhen(true)] out IEnumerable<ReadCellResult>? result)
    {
        var cells = new List<ReadCellResult>();
        foreach (var cellReader in Readers)
        {
            if (cellReader.TryGetValue(reader, preserveFormatting, out var cellValues))
            {
                cells.AddRange(cellValues);
            }
        }

        if (cells.Count > 0)
        {
            result = cells;
            return true;
        }
        else
        {
            result = null;
            return false;
        }
    }
}
