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

    private IExcelDataReader? _reader;
    private bool _preserveFormatting;
    private int _readerIndex;
    private ReadCellResult? _current;

    /// <summary>
    /// Initializes a new instance of <see cref="CompositeCellsReader"/>.
    /// </summary>
    /// <param name="readers">The cell readers.</param>
    /// <exception cref="ArgumentException">Thrown when the readers list is empty or contains null values.</exception>
    public CompositeCellsReader(params IReadOnlyList<ICellReader> readers)
    {
        ThrowHelpers.ThrowIfNull(readers, nameof(readers));
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
    public bool Start(IExcelDataReader reader, bool preserveFormatting, out int count)
    {
        _reader = reader;
        _preserveFormatting = preserveFormatting;
        _readerIndex = -1;
        _current = null;

        // Count how many readers successfully produce values
        int resultCount = 0;
        for (int i = 0; i < Readers.Count; i++)
        {
            if (Readers[i].TryGetValue(reader, preserveFormatting, out _))
            {
                resultCount++;
            }
        }

        count = resultCount;
        return true;
    }

    /// <inheritdoc/>
    public bool TryGetNext([NotNullWhen(true)] out ReadCellResult result)
    {
        while (true)
        {
            _readerIndex++;
            if (_readerIndex >= Readers.Count)
            {
                result = default;
                return false;
            }

            if (Readers[_readerIndex].TryGetValue(_reader!, _preserveFormatting, out var readResult))
            {
                _current = readResult;
                result = readResult;
                return true;
            }
            // Continue to the next reader
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        _readerIndex = -1;
        _current = null;
        _reader = null;
    }
}
