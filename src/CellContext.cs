using System.Globalization;
using System.Reflection;

namespace ExcelMapper;

/// <summary>
/// Where a cell being converted is: its sheet, row and column, the heading of the rows being read,
/// and the member it's mapped to. See <see cref="ICellConverter{T}"/>.
/// </summary>
public readonly struct CellContext
{
    // The view the rows are read with: for a region, the region's, with its heading.
    private readonly SheetView _view;

    /// <summary>
    /// Creates a context, for example to test a converter.
    /// </summary>
    public CellContext(ExcelSheet sheet, int rowIndex, int columnIndex, MemberInfo? member)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        ThrowHelpers.ThrowIfNegative(rowIndex, nameof(rowIndex));
        ThrowHelpers.ThrowIfNegative(columnIndex, nameof(columnIndex));

        _view = sheet.View;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Member = member;
        _options = null;
    }

    /// <summary>
    /// Creates the context of a cell being mapped.
    /// </summary>
    internal CellContext(SheetView sheet, int rowIndex, int columnIndex, MemberInfo? member, ConversionOptions? options)
    {
        _view = sheet;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Member = member;
        _options = options;
    }

    private readonly ConversionOptions? _options;

    /// <summary>
    /// Gets how the member's built-in conversion parses text (its formats, format provider and so
    /// on), from the pipeline converting the cell.
    /// </summary>
    internal ConversionOptions Options => _options ?? ConversionOptions.Default;

    /// <summary>
    /// Gets the sheet being read. When reading a region, this is the sheet the region was created
    /// from: use <see cref="Heading"/> for the region's heading.
    /// </summary>
    public ExcelSheet Sheet => _view?.Sheet!;

    /// <summary>
    /// Gets the heading of the rows being read: the region's heading when reading a region,
    /// otherwise the sheet's. Null if they have no heading.
    /// </summary>
    public ExcelHeading? Heading => _view?.Heading;

    /// <summary>
    /// Gets the sheet the rows are read with, which has the heading and columns of a region.
    /// </summary>
    internal SheetView View => _view;

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
    /// <see cref="ExcelImporterConfiguration.Culture"/> (by default <see cref="CultureInfo.InvariantCulture"/>).
    /// </summary>
    public IFormatProvider FormatProvider => _view?.Culture ?? CultureInfo.InvariantCulture;

    /// <summary>
    /// Gets the cell's address, such as "B11".
    /// </summary>
    public string CellAddress => ExcelRange.GetCellAddress(RowIndex, ColumnIndex);
}
