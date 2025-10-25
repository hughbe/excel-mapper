using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a cell and produces multiple values by splitting the string value
/// using the given separators.
/// </summary>
public abstract class SplitReaderFactory : ICellsReaderFactory, IColumnIndicesProviderCellReaderFactory, IColumnNamesProviderCellReaderFactory
{
    /// <summary>
    /// Gets or sets the options used to split the string value of the cell.
    /// </summary>
    public StringSplitOptions Options { get; set; }

    private ICellReaderFactory _readerFactory;

    /// <summary>
    /// Gets or sets the ICellReader that reads the string value of the cell
    /// before it is split.
    /// </summary>
    public ICellReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            _readerFactory = value;
        }
    }

    /// <summary>
    /// Constructs a reader that reads the string value of a cell and produces multiple values
    /// by splitting it.
    /// </summary>
    /// <param name="readerFactory">The ICellReaderFactory that reads the string value of the cell before it is split.</param>
    public SplitReaderFactory(ICellReaderFactory readerFactory)
    {
        ArgumentNullException.ThrowIfNull(readerFactory);
        _readerFactory = readerFactory;
    }

    /// <summary>
    /// Splits the given string value into multiple values.
    /// </summary>
    /// <param name="value">The string value to split.</param>
    /// <returns>The multiple values produced by splitting the string value.</returns>
    protected abstract string[] GetValues(string value);

    /// <inheritdoc/>
    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ICellReader? reader = _readerFactory.GetCellReader(sheet);
        if (reader == null)
        {
            return null;
        }

        return new Reader(reader, this);
    }

    /// <inheritdoc/>
    public IReadOnlyList<string>? GetColumnNames(ExcelSheet sheet)
    {
        if (_readerFactory is IColumnNameProviderCellReaderFactory nameProvider)
        {
            var columnName = nameProvider.GetColumnName(sheet);
            if (!string.IsNullOrEmpty(columnName))
            {
                return [columnName];
            }
        }

        return null;
    }

    /// <inheritdoc/>
    public IReadOnlyList<int>? GetColumnIndices(ExcelSheet sheet)
    {
        if (_readerFactory is IColumnIndexProviderCellReaderFactory indexProvider)
        {
            var columnIndex = indexProvider.GetColumnIndex(sheet);
            if (columnIndex != null && columnIndex != -1)
            {
                return [columnIndex.Value];
            }
        }

        return null;
    }

    private class Reader(ICellReader Reader, SplitReaderFactory Splitter) : ICellsReader
    {
        public bool TryGetValues(IExcelDataReader reader, bool preserveFormatting, [NotNullWhen(true)] out IEnumerable<ReadCellResult>? result)
        {
            if (!Reader.TryGetValue(reader, preserveFormatting, out var readResult))
            {
                result = null;
                return false;
            }

            var stringValue = readResult.GetString();
            if (stringValue == null)
            {
                result = [];
                return true;
            }

            result = Splitter
                .GetValues(stringValue)
                .Select(s => new ReadCellResult(readResult.ColumnIndex, s, preserveFormatting));
            return true;
        }
    }
}
