using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper;

/// <summary>
/// An object that represents a single sheet of an excel document.
/// </summary>
/// <remarks>
/// <para>
/// This class maintains mutable state (such as the current row index) and is not thread-safe.
/// Each <see cref="ExcelSheet"/> instance should be used by only one thread at a time.
/// </para>
/// <para>
/// If you need to process rows concurrently, read all rows into a collection first using
/// <see cref="ReadRows{T}()"/> and then process the collection in parallel.
/// </para>
/// </remarks>
public class ExcelSheet
{
    /// <summary>
    /// Creates a sheet for a worksheet.
    /// </summary>
    internal ExcelSheet(Worksheet worksheet)
    {
        View = new SheetView(worksheet, this);
    }

    /// <summary>
    /// Gets what the sheet reads, and its position: the reading itself is done by the view, which
    /// a region also reads through (see <see cref="ExcelRegion"/>).
    /// </summary>
    internal SheetView View { get; }

    internal Worksheet Worksheet => View.Worksheet;

    /// <summary>
    /// Gets the name of the sheet.
    /// </summary>
    public string Name => Worksheet.Name;

    /// <summary>
    /// Gets the visibility of the sheet.
    /// </summary>
    public ExcelSheetVisibility Visibility => Worksheet.Visibility;

    /// <summary>
    /// Gets the zero-based index of the sheet where 0 is the first sheet in the document.
    /// </summary>
    public int Index => Worksheet.Index;

    /// <summary>
    /// Gets the number of columns in the sheet.
    /// </summary>
    public int NumberOfColumns => Worksheet.NumberOfColumns;

    /// <summary>
    /// Gets or sets whether the sheet has a heading. This is true by default.
    /// </summary>
    [MemberNotNullWhen(true, nameof(Heading))]
    public bool HasHeading
    {
        get => View.HasHeading;
        set => View.HasHeading = value;
    }

    /// <summary>
    /// Gets or sets the zero-based index of row containing the heading. This is 0 (the first row) by default.
    /// If the value is non-zero, all rows preceding the heading are skipped and not mapped.
    /// </summary>
    public int HeadingIndex
    {
        get => View.HeadingIndex;
        set => View.HeadingIndex = value;
    }

    /// <summary>
    /// Gets a region of the whole worksheet, with the first row as its heading. Unlike reading the
    /// sheet, reading a region always starts at its top, and doesn't use or change this sheet's
    /// heading, <see cref="DataRange"/> or position.
    /// </summary>
    /// <returns>The region.</returns>
    public ExcelRegion Region() => new(this, default);

    /// <summary>
    /// Gets a region of the worksheet, with its first row as its heading. Use regions to read part of
    /// a sheet, or several tables on one sheet.
    /// </summary>
    /// <param name="range">The rows and columns of the region, in worksheet coordinates.</param>
    /// <returns>The region.</returns>
    public ExcelRegion Region(ExcelRange range) => new(this, range);

    /// <summary>
    /// Gets a region of the worksheet, with its first row as its heading. Use regions to read part of
    /// a sheet, or several tables on one sheet.
    /// </summary>
    /// <param name="address">The address of the region, such as "B3:F20", "$B$3:$F$20", "B:F" or "3:20".</param>
    /// <returns>The region.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="address"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="address"/> is empty or not a valid address.</exception>
    public ExcelRegion Region(string address) => new(this, ExcelRange.Parse(address));

    /// <summary>
    /// Gets or sets the range of rows and columns that contain data to be mapped. By default this is all rows and columns.
    /// </summary>
    public ExcelRange DataRange
    {
        get => View.DataRange;
        set => View.DataRange = value;
    }

    /// <summary>
    /// Gets the heading that was read from the sheet. This will return null if HasHeading is false
    /// or the heading has not been read yet by calling ReadHeading or ReadRows.
    /// </summary>
    public ExcelHeading? Heading => View.Heading;

    /// <summary>
    /// Gets the index of the row currently being mapped.
    /// </summary>
    public int CurrentRowIndex => View.CurrentRowIndex;

    /// <summary>
    /// Reads the heading of the sheet including column names and indices.
    /// </summary>
    /// <returns>An object that represents the heading of the sheet.</returns>
    public ExcelHeading ReadHeading() => View.ReadHeading();

    /// <summary>
    /// Maps each row of the sheet to an object using a registered mapping. If no map is registered for this
    /// type then the type will be automapped. This method will read the sheet's heading if the sheet has
    /// a heading and the heading has not yet been read.
    /// </summary>
    /// <remarks>
    /// Reading is forward-only: rows are read from the current position of the sheet. Rows that have
    /// already been read (for example by a previous call to this method or <see cref="ReadRow{T}()"/>)
    /// are not returned again, so calling this method after all rows have been read returns an empty
    /// sequence. Use <see cref="ReadRows{T}(int, int)"/> to re-read a specific range of rows.
    /// </remarks>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <returns>A list of objects of type T mapped from each row in the sheet.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => View.ReadRows<T>();

    /// <summary>
    /// Maps each row of the sheet to an object with <paramref name="classMap"/>. See
    /// <see cref="ReadRows{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <returns>The objects mapped from each row.</returns>
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap) => View.ReadRows(SheetView.Given(classMap));

    /// <summary>
    /// Maps each row within the range specified to an object using a registered mapping. If no map is registered for this
    /// type then the type will be automapped. This method will read the sheet's heading if the sheet has
    /// a heading and the heading has not yet been read. If rows beyond <paramref name="startIndex"/> have
    /// already been read, the sheet is rewound so the requested range is returned.
    /// </summary>
    /// <param name="startIndex">The zero-based index from the first row of the document (including the header) of the range of rows to map from.</param>
    /// <param name="count">The number of rows to read and map.</param>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <returns>A list of objects of type T mapped from each row within the range specified.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int startIndex, int count) => View.ReadRows<T>(startIndex, count);

    /// <summary>
    /// Maps each row within the range specified to an object with <paramref name="classMap"/>. See
    /// <see cref="ReadRows{T}(int, int)"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <param name="startIndex">The zero-based index from the first row of the document (including the header) of the range of rows to map from.</param>
    /// <param name="count">The number of rows to read and map.</param>
    /// <returns>The objects mapped from each row within the range specified.</returns>
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap, int startIndex, int count) => View.ReadRows(SheetView.Given(classMap), startIndex, count);

    /// <summary>
    /// Maps a single row of a sheet to an object using a registered mapping. If no map is registered for this
    /// type then the type will be automapped. This method will read the sheet's heading if the sheet has a
    /// heading and the heading has not yet been read. This method will throw if mapping fails or there are
    /// no more rows left.
    /// </summary>
    /// <typeparam name="T">The type of the object to map a single row to.</typeparam>
    /// <returns>An object of type T mapped from a single row in the sheet.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public T ReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => View.ReadRow<T>();

    /// <summary>
    /// Maps a single row of the sheet to an object with <paramref name="classMap"/>. See
    /// <see cref="ReadRow{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map a single row to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <returns>The object mapped from the row.</returns>
    public T ReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap) => View.ReadRow(SheetView.Given(classMap));

    /// <summary>
    /// Maps a single row of a sheet to an object using a registered mapping. If no map is registered for this
    /// type then the type will be automapped. This method will read the sheet's heading if the sheet has a
    /// heading and the heading has not yet been read.
    /// </summary>
    /// <typeparam name="T">The type of the object to map a single row to.</typeparam>
    /// <param name="value">An object of type T mapped from a single row in the sheet.</param>
    /// <returns>False if there are no more rows in the sheet or the row cannot be mapped to an object, else true.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>([NotNullWhen(true)] out T? value) => View.TryReadRow(out value);

    /// <summary>
    /// Maps a single row of the sheet to an object with <paramref name="classMap"/>. See
    /// <see cref="TryReadRow{T}(out T)"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map a single row to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <param name="value">The object mapped from the row.</param>
    /// <returns>False if there are no more rows in the sheet or the row cannot be mapped to an object, else true.</returns>
    public bool TryReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap, [NotNullWhen(true)] out T? value) => View.TryReadRow(SheetView.Given(classMap), out value);

    /// <summary>
    /// Maps each row of the sheet to an object, like <see cref="ReadRows{T}()"/>, but reports invalid
    /// rows instead of throwing at the first invalid cell. Every cell that can't be mapped is
    /// recorded, so a single pass over a sheet can report all of its problems.
    /// </summary>
    /// <remarks>
    /// Problems with the mapping itself, such as a type that can't be auto-mapped, still throw.
    /// Use <see cref="ExcelRowResult{T}.IsValid"/> to find the rows that were mapped successfully.
    /// </remarks>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <returns>The result of mapping each row in the sheet.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<ExcelRowResult<T>> ReadRowResults<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => View.ReadRowResults<T>();

    /// <summary>
    /// Maps each row of the sheet with <paramref name="classMap"/>, reporting invalid rows instead
    /// of throwing at the first invalid cell. See <see cref="ReadRowResults{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <returns>The result of mapping each row in the sheet.</returns>
    public IEnumerable<ExcelRowResult<T>> ReadRowResults<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap) => View.ReadRowResults(SheetView.Given(classMap));

    /// <summary>
    /// Checks that the sheet has every column needed to map rows to <typeparamref name="T"/>, and
    /// throws an <see cref="ExcelMappingException"/> listing all of the missing columns if not.
    /// Call this before reading rows to report every missing column at once, rather than one at a
    /// time when the first row is read. Optional members are not required. This reads the heading
    /// if it has not been read yet.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public void ValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => View.ValidateHeading<T>();

    /// <summary>
    /// Checks that the sheet has every column <paramref name="classMap"/> needs, and throws an
    /// <see cref="ExcelMappingException"/> listing all of the missing columns if not. See
    /// <see cref="ValidateHeading{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    public void ValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap) => View.ValidateHeading(SheetView.Given(classMap));

    /// <summary>
    /// Checks that the sheet has every column needed to map rows to <typeparamref name="T"/>.
    /// Optional members are not required. This reads the heading if it has not been read yet.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <param name="problems">A description of each missing column, such as
    /// <c>column "Salary" (member Salary)</c>. Empty if none are missing.</param>
    /// <returns>True if every required column is present, else false.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(out IReadOnlyList<string> problems) => View.TryValidateHeading<T>(out problems);

    /// <summary>
    /// Checks that the sheet has every column <paramref name="classMap"/> needs. See
    /// <see cref="TryValidateHeading{T}(out IReadOnlyList{string})"/>.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <param name="problems">A description of each missing column. Empty if none are missing.</param>
    /// <returns>True if every required column is present, else false.</returns>
    public bool TryValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap, out IReadOnlyList<string> problems) => View.TryValidateHeading(SheetView.Given(classMap), out problems);

    /// <summary>
    /// Finds the heading of a sheet with rows above it, such as a title or notes, and reads it. The
    /// heading is the first row, among the first 50 rows, that
    /// has every column <typeparamref name="T"/> needs by name and at least one of its columns.
    /// Blank rows are skipped. Rows are then read from the row after the heading.
    /// </summary>
    /// <remarks>
    /// The search starts at <see cref="HeadingIndex"/> (the first row by default), and sets
    /// <see cref="HeadingIndex"/> and <see cref="HasHeading"/>. Columns found by index or letter,
    /// optional columns and dictionaries of all columns don't identify the heading row, but optional
    /// columns help choose between rows.
    /// </remarks>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <returns>The heading.</returns>
    /// <exception cref="ExcelMappingException">No row has the columns <typeparamref name="T"/> needs; the message
    /// gives the closest row and its missing columns. Or <typeparamref name="T"/> doesn't look up any column by name.</exception>
    /// <exception cref="InvalidOperationException">The heading has already been read.</exception>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public ExcelHeading FindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => View.FindHeading<T>(HeadingSearch.DefaultMaxRows);

    /// <summary>
    /// Finds the heading of a sheet with rows above it, among the first 50 rows, from the columns
    /// <paramref name="classMap"/> needs, and reads it. See <see cref="FindHeading{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <returns>The heading.</returns>
    /// <exception cref="ExcelMappingException">No row has the columns the class map needs.</exception>
    /// <exception cref="InvalidOperationException">The heading has already been read.</exception>
    public ExcelHeading FindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap) => View.FindHeading(SheetView.Given(classMap), HeadingSearch.DefaultMaxRows);

    /// <summary>
    /// Finds the heading of a sheet with rows above it, such as a title or notes, and reads it. The
    /// heading is the first row, among the first <paramref name="maxRows"/> rows, that has every
    /// column <typeparamref name="T"/> needs by name and at least one of its columns.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <param name="maxRows">The number of rows to search, starting at <see cref="HeadingIndex"/>.</param>
    /// <returns>The heading.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="maxRows"/> is zero or negative.</exception>
    /// <exception cref="ExcelMappingException">No row has the columns <typeparamref name="T"/> needs; the message
    /// gives the closest row and its missing columns. Or <typeparamref name="T"/> doesn't look up any column by name.</exception>
    /// <exception cref="InvalidOperationException">The heading has already been read.</exception>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public ExcelHeading FindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int maxRows) => View.FindHeading<T>(maxRows);

    /// <summary>
    /// Finds the heading of a sheet with rows above it, among the first <paramref name="maxRows"/>
    /// rows, from the columns <paramref name="classMap"/> needs, and reads it. See
    /// <see cref="FindHeading{T}(int)"/>.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <param name="maxRows">The number of rows to search, starting at <see cref="HeadingIndex"/>.</param>
    /// <returns>The heading.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="maxRows"/> is zero or negative.</exception>
    /// <exception cref="ExcelMappingException">No row has the columns the class map needs.</exception>
    /// <exception cref="InvalidOperationException">The heading has already been read.</exception>
    public ExcelHeading FindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap, int maxRows) => View.FindHeading(SheetView.Given(classMap), maxRows);

    /// <summary>
    /// Tries to find the heading of a sheet with rows above it, such as a title or notes, and reads it.
    /// See <see cref="FindHeading{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <param name="heading">The heading, if found.</param>
    /// <returns>Whether a heading was found. If not, the sheet is unchanged and reading starts from the beginning again.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryFindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>([NotNullWhen(true)] out ExcelHeading? heading)
        => View.TryFindHeading<T>(HeadingSearch.DefaultMaxRows, out heading, out _);

    /// <summary>
    /// Tries to find the heading of a sheet with rows above it from the columns
    /// <paramref name="classMap"/> needs, and reads it. See <see cref="FindHeading{T}(ExcelClassMap{T})"/>.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <param name="heading">The heading, if found.</param>
    /// <returns>Whether a heading was found. If not, the sheet is unchanged and reading starts from the beginning again.</returns>
    public bool TryFindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap, [NotNullWhen(true)] out ExcelHeading? heading)
        => View.TryFindHeading(SheetView.Given(classMap), HeadingSearch.DefaultMaxRows, out heading, out _);

    /// <summary>
    /// Tries to find the heading of a sheet with rows above it, such as a title or notes, and reads it.
    /// See <see cref="FindHeading{T}(int)"/>.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <param name="maxRows">The number of rows to search, starting at <see cref="HeadingIndex"/>.</param>
    /// <param name="heading">The heading, if found.</param>
    /// <returns>Whether a heading was found. If not, the sheet is unchanged and reading starts from the beginning again.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryFindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int maxRows, [NotNullWhen(true)] out ExcelHeading? heading)
        => View.TryFindHeading<T>(maxRows, out heading, out _);

    /// <summary>
    /// Tries to find the heading of a sheet with rows above it from the columns
    /// <paramref name="classMap"/> needs, and reads it. See <see cref="FindHeading{T}(ExcelClassMap{T}, int)"/>.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    /// <param name="classMap">The class map to map rows with, such as one from an
    /// <see cref="ExcelMappingContext"/>. Unlike the overload without one, this doesn't look up or
    /// auto-map a class map, so it works with native AOT.</param>
    /// <param name="maxRows">The number of rows to search, starting at <see cref="HeadingIndex"/>.</param>
    /// <param name="heading">The heading, if found.</param>
    /// <returns>Whether a heading was found. If not, the sheet is unchanged and reading starts from the beginning again.</returns>
    public bool TryFindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap, int maxRows, [NotNullWhen(true)] out ExcelHeading? heading)
        => View.TryFindHeading(SheetView.Given(classMap), maxRows, out heading, out _);
}
