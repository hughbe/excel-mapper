using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper;

/// <summary>
/// A block of rows and columns of a worksheet, read with its own heading. Use it to read part of a
/// sheet, or several tables on one sheet. Create one with <see cref="ExcelSheet.Region()"/>.
/// </summary>
/// <remarks>
/// <para>
/// A region is immutable: the <c>With</c> methods return a new region. Each read starts at the top
/// of the region, so a region can be read any number of times, and reading it doesn't change the
/// <see cref="ExcelSheet"/> it was created from (its heading, <see cref="ExcelSheet.DataRange"/> or
/// position) or other regions.
/// </para>
/// <para>
/// Columns chosen by position (such as <see cref="ExcelColumnIndexAttribute"/> or
/// <see cref="ExcelColumnLetterAttribute"/>) count from the region's first column, so one type can
/// map tables of the same shape anywhere on a sheet. Addresses and error messages use the sheet's
/// rows and columns.
/// </para>
/// <para>
/// The importer reads the file forward only. Reading one region at a time is as fast as reading the
/// sheet; alternating between regions row by row re-reads the file each time.
/// </para>
/// </remarks>
public sealed class ExcelRegion
{
    private enum HeadingKind
    {
        FirstRow,
        AtRow,
        None,
        Search,
    }

    private readonly HeadingKind _headingKind;
    private readonly int _headingValue;

    internal ExcelRegion(ExcelSheet sheet, ExcelRange range)
        : this(sheet, range, HeadingKind.FirstRow, 0)
    {
    }

    private ExcelRegion(ExcelSheet sheet, ExcelRange range, HeadingKind headingKind, int headingValue)
    {
        Sheet = sheet;
        Range = range;
        _headingKind = headingKind;
        _headingValue = headingValue;
    }

    /// <summary>
    /// Gets the sheet the region is part of.
    /// </summary>
    public ExcelSheet Sheet { get; }

    /// <summary>
    /// Gets the rows and columns of the region, in worksheet coordinates.
    /// </summary>
    public ExcelRange Range { get; }

    /// <summary>
    /// Returns a region whose heading is the given row of the worksheet, inside or above the region.
    /// The data is the region's rows after the heading.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the heading row in the worksheet.</param>
    /// <returns>A new region.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="rowIndex"/> is negative.</exception>
    public ExcelRegion WithHeadingAt(int rowIndex)
    {
        ThrowHelpers.ThrowIfNegative(rowIndex, nameof(rowIndex));
        return new ExcelRegion(Sheet, Range, HeadingKind.AtRow, rowIndex);
    }

    /// <summary>
    /// Returns a region without a heading: every row is data, and columns are found by index or letter.
    /// </summary>
    /// <returns>A new region.</returns>
    public ExcelRegion WithoutHeading() => new(Sheet, Range, HeadingKind.None, 0);

    /// <summary>
    /// Returns a region whose heading is found when it's read: the first row, among the region's first
    /// 50 rows, that has every column the mapped type needs by name and at least one of its columns.
    /// See <see cref="ExcelSheet.FindHeading{T}()"/>.
    /// </summary>
    /// <returns>A new region.</returns>
    public ExcelRegion WithHeadingSearch() => new(Sheet, Range, HeadingKind.Search, HeadingSearch.DefaultMaxRows);

    /// <summary>
    /// Returns a region whose heading is found when it's read: the first row, among the region's first
    /// <paramref name="maxRows"/> rows, that has every column the mapped type needs by name and at least
    /// one of its columns. See <see cref="ExcelSheet.FindHeading{T}(int)"/>.
    /// </summary>
    /// <param name="maxRows">The number of rows to search.</param>
    /// <returns>A new region.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="maxRows"/> is zero or negative.</exception>
    public ExcelRegion WithHeadingSearch(int maxRows)
    {
        ThrowHelpers.ThrowIfNegativeOrZero(maxRows, nameof(maxRows));
        return new ExcelRegion(Sheet, Range, HeadingKind.Search, maxRows);
    }

    /// <summary>
    /// Maps each row of the region to an object. Reading starts at the top of the region each time the
    /// result is enumerated.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <returns>The objects mapped from each row.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => ReadRows(SheetView.AutoMapped<T>.ClassMap);

    /// <summary>
    /// Maps each row of the region to an object with <paramref name="classMap"/>. See
    /// <see cref="ReadRows{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <returns>The objects mapped from each row.</returns>
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap) => ReadRows(SheetView.Given(classMap));

    private IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap)
    {
        foreach (var row in CreateView(classMap).ReadRows(classMap))
        {
            yield return row;
        }
    }

    /// <summary>
    /// Maps each row of the region, reporting every invalid cell instead of stopping at the first.
    /// See <see cref="ExcelSheet.ReadRowResults{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <returns>The result of mapping each row.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<ExcelRowResult<T>> ReadRowResults<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => ReadRowResults(SheetView.AutoMapped<T>.ClassMap);

    /// <summary>
    /// Maps each row of the region with <paramref name="classMap"/>, reporting every invalid cell
    /// instead of stopping at the first. See <see cref="ReadRowResults{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <returns>The result of mapping each row.</returns>
    public IEnumerable<ExcelRowResult<T>> ReadRowResults<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap) => ReadRowResults(SheetView.Given(classMap));

    private IEnumerable<ExcelRowResult<T>> ReadRowResults<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap)
    {
        foreach (var result in CreateView(classMap).ReadRowResults(classMap))
        {
            yield return result;
        }
    }

    /// <summary>
    /// Maps each row of the region, reporting the problems in invalid rows to
    /// <paramref name="errors"/> and skipping those rows. See
    /// <see cref="ExcelSheet.ReadRows{T}(IExcelErrorSink)"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <param name="errors">Receives each problem as it's found.</param>
    /// <returns>The objects mapped from the rows without problems.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="errors"/> is null.</exception>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(IExcelErrorSink errors) => ReadRows(SheetView.AutoMapped<T>.ClassMap, errors);

    /// <summary>
    /// Maps each row of the region with <paramref name="classMap"/>, reporting the problems in
    /// invalid rows to <paramref name="errors"/> and skipping those rows. See
    /// <see cref="ExcelSheet.ReadRows{T}(IExcelErrorSink)"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <param name="errors">Receives each problem as it's found.</param>
    /// <returns>The objects mapped from the rows without problems.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="classMap"/> or <paramref name="errors"/> is null.</exception>
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap, IExcelErrorSink errors) => ReadRows(SheetView.Given(classMap), errors);

    private IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, IExcelErrorSink errors)
    {
        ThrowHelpers.ThrowIfNull(errors, nameof(errors));
        return ReadRowsWithSink(classMap, errors);
    }

    private IEnumerable<T> ReadRowsWithSink<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, IExcelErrorSink errors)
    {
        foreach (var row in CreateView(classMap).ReadRows(classMap, errors))
        {
            yield return row;
        }
    }

    /// <summary>
    /// Checks that the region's heading has every column that <typeparamref name="T"/> needs.
    /// See <see cref="ExcelSheet.ValidateHeading{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type rows will be mapped to.</typeparam>
    /// <exception cref="ExcelMappingException">A column is missing, or the heading can't be found.</exception>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public void ValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
        => ValidateHeading(SheetView.AutoMapped<T>.ClassMap);

    /// <summary>
    /// Checks that the region's heading has every column <paramref name="classMap"/> needs. See
    /// <see cref="ValidateHeading{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type rows will be mapped to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <exception cref="ExcelMappingException">A column is missing, or the heading can't be found.</exception>
    public void ValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap) => ValidateHeading(SheetView.Given(classMap));

    private void ValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap)
        => CreateView(classMap).ValidateHeading(classMap);

    /// <summary>
    /// Checks that the region's heading has every column that <typeparamref name="T"/> needs.
    /// See <see cref="ExcelSheet.TryValidateHeading{T}(out IReadOnlyList{string})"/>.
    /// </summary>
    /// <typeparam name="T">The type rows will be mapped to.</typeparam>
    /// <param name="problems">The missing columns, or why the heading couldn't be found.</param>
    /// <returns>Whether every column is present.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(out IReadOnlyList<string> problems)
        => TryValidateHeading(SheetView.AutoMapped<T>.ClassMap, out problems);

    /// <summary>
    /// Checks that the region's heading has every column <paramref name="classMap"/> needs. See
    /// <see cref="TryValidateHeading{T}(out IReadOnlyList{string})"/>.
    /// </summary>
    /// <typeparam name="T">The type rows will be mapped to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <param name="problems">The missing columns, or why the heading couldn't be found.</param>
    /// <returns>Whether every column is present.</returns>
    public bool TryValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap, out IReadOnlyList<string> problems)
        => TryValidateHeading(SheetView.Given(classMap), out problems);

    private bool TryValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, out IReadOnlyList<string> problems)
    {
        var view = CreateView();
        if (_headingKind == HeadingKind.Search && !view.TryFindHeading(classMap, _headingValue, out _, out var error))
        {
            problems = [error];
            return false;
        }

        return view.TryValidateHeading(classMap, out problems);
    }

    private SheetView CreateView<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap)
    {
        var view = CreateView();
        if (_headingKind == HeadingKind.Search)
        {
            view.FindHeading(classMap, _headingValue);
        }

        return view;
    }

    /// <summary>
    /// Creates a view of the region's worksheet, with its own heading and position, set up to read
    /// the region. The view's sheet is the one the region is of, which extension points and
    /// exceptions are given.
    /// </summary>
    private SheetView CreateView()
    {
        var view = new SheetView(Sheet.Worksheet, Sheet);

        // Columns chosen by position count from the region's first column, and stop at its last.
        var columnCount = Sheet.NumberOfColumns;
        view.ColumnOffset = Math.Min(Range.Columns.Start.GetOffset(columnCount), columnCount);
        view.ColumnEnd = Math.Max(view.ColumnOffset, Math.Min(Range.Columns.End.GetOffset(columnCount), columnCount));

        switch (_headingKind)
        {
            case HeadingKind.None:
                view.HasHeading = false;
                view.DataRange = Range;
                break;

            case HeadingKind.AtRow:
                // The heading is read from its row; the data is the region's rows after it.
                view.DataRange = new ExcelRange(_headingValue..Range.Rows.End, Range.Columns);
                view.FirstDataRow = Range.Rows.Start;
                break;

            default:
                view.DataRange = Range;
                break;
        }

        return view;
    }
}
