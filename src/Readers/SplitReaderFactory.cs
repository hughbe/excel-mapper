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

    protected abstract string[] GetValues(string value);

    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ICellReader? reader = _readerFactory.GetCellReader(sheet);
        if (reader == null)
        {
            return null;
        }

        return new Reader(reader, this);
    }

    public string[]? GetColumnNames(ExcelSheet sheet)
    {
        if (_readerFactory is IColumnNameProviderCellReaderFactory nameProvider)
        {
            return [nameProvider.GetColumnName(sheet)];
        }

        return null;
    }


    public int[]? GetColumnIndices(ExcelSheet sheet)
    {
        if (_readerFactory is IColumnIndexProviderCellReaderFactory indexProvider)
        {
            return [indexProvider.GetColumnIndex(sheet)];
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
