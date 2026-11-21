using System.Globalization;
using System.Reflection;

namespace ExcelMapper;

/// <summary>
/// Where a cell being converted is: its sheet, row and column, and the member it's mapped to. See
/// <see cref="ICellConverter{T}"/>.
/// </summary>
public readonly struct CellContext
{
    /// <summary>
    /// Creates a context, for example to test a converter.
    /// </summary>
    public CellContext(ExcelSheet sheet, int rowIndex, int columnIndex, MemberInfo? member)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        ThrowHelpers.ThrowIfNegative(rowIndex, nameof(rowIndex));
        ThrowHelpers.ThrowIfNegative(columnIndex, nameof(columnIndex));

        Sheet = sheet;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Member = member;
    }

    /// <summary>
    /// Gets the sheet being read.
    /// </summary>
    public ExcelSheet Sheet { get; }

    /// <summary>
    /// Gets the zero-based index of the cell's row in the worksheet.
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// Gets the zero-based index of the cell's column in the worksheet.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// Gets the property or field the value is for, or null if there isn't one, for example when a
    /// row is read as a primitive with <c>ReadRows&lt;int&gt;()</c>.
    /// </summary>
    public MemberInfo? Member { get; }

    /// <summary>
    /// Gets the culture to parse and format values with: the importer's
    /// <see cref="ExcelImporterConfiguration.Culture"/>, or the current culture if it isn't set.
    /// </summary>
    public IFormatProvider FormatProvider => Sheet?.Culture ?? CultureInfo.CurrentCulture;

    /// <summary>
    /// Gets the cell's address, such as "B11".
    /// </summary>
    public string CellAddress => ExcelRange.GetCellAddress(RowIndex, ColumnIndex);
}
