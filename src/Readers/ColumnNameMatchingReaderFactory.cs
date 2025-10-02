using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a single cell given the predicate matching the column name.
/// </summary>
public sealed class ColumnNameMatchingReaderFactory : ICellReaderFactory
{
    private readonly Func<string, bool> _predicate;

    /// <summary>
    /// Constructs a reader that reads the value of a single cell given the predicate matching the column name.
    /// </summary>
    /// <param name="predicate">The predicate containing the column name to read.</param>
    public ColumnNameMatchingReaderFactory(Func<string, bool> predicate)
    {
        _predicate = predicate ?? throw new ArgumentNullException(nameof(predicate));
    }

    public ICellReader? GetReader(ExcelSheet sheet)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        if (!sheet.Heading.TryGetFirstColumnMatchingIndex(_predicate, out int columnIndex))
        {
            return null;
        }

        return new ColumnIndexReader(columnIndex);
    }
}