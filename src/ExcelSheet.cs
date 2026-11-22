using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ExcelDataReader;
using ExcelMapper.Readers;

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
    /// Creates a sheet for a worksheet. A region is read with a sheet of its own for the same
    /// worksheet, with its own heading and position (see <see cref="ExcelRegion"/>).
    /// </summary>
    internal ExcelSheet(Worksheet worksheet)
    {
        Worksheet = worksheet;
        _cursor = new RowCursor(worksheet);
        _bindings = new SheetBindings(this);
    }

    /// <summary>
    /// The worksheet, shared with the sheets that regions of it are read with.
    /// </summary>
    internal Worksheet Worksheet { get; }

    private readonly RowCursor _cursor;
    private readonly SheetBindings _bindings;

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

    private bool _hasHeading = true;

    /// <summary>
    /// Gets or sets whether the sheet has a heading. This is true by default.
    /// </summary>
    [MemberNotNullWhen(true, nameof(Heading))]
    public bool HasHeading
    {
        get => _hasHeading;
        set
        {
            if (Heading != null)
            {
                throw new InvalidOperationException("The heading has already been read. Set this property before reading any rows.");
            }

            _hasHeading = value;
        }
    }

    /// <summary>
    /// Gets or sets the zero-based index of row containing the heading. This is 0 (the first row) by default.
    /// If the value is non-zero, all rows preceding the heading are skipped and not mapped.
    /// </summary>
    public int HeadingIndex
    {
        get => _dataRange.Rows.Start.Value;
        set
        {
            ThrowHelpers.ThrowIfNegative(value, nameof(value));
            if (!HasHeading)
            {
                throw new InvalidOperationException("The sheet has no heading.");
            }
            if (Heading != null)
            {
                throw new InvalidOperationException("The heading has already been read.");
            }

            // Adjust the data range. The data range starts at the heading index.
            var rowsRange = value.._dataRange.Rows.End;
            _dataRange = new ExcelRange(rowsRange, _dataRange.Columns);
        }
    }

    private ExcelRange _dataRange = new();

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
        get => _dataRange;
        set
        {
            if (Heading != null)
            {
                throw new InvalidOperationException("The heading has already been read. Set this property before reading any rows.");
            }

            // Set the data range.
            _dataRange = value;
        }
    }

    /// <summary>
    /// Gets the heading that was read from the sheet. This will return null if HasHeading is false
    /// or the heading has not been read yet by calling ReadHeading or ReadRows.
    /// </summary>
    public ExcelHeading? Heading { get; internal set; }

    /// <summary>
    /// Gets the index of the row currently being mapped.
    /// </summary>
    public int CurrentRowIndex => _cursor.RowIndex;

    /// <summary>
    /// The importer's culture, if set (see <see cref="ExcelImporterConfiguration.Culture"/>).
    /// </summary>
    internal CultureInfo? Culture => Worksheet.Configuration.Culture;

    internal ExcelImporterConfiguration Configuration => Worksheet.Configuration;

    private IExcelDataReader Reader => Worksheet.Reader;

    /// <summary>
    /// Reads the heading of the sheet including column names and indices.
    /// </summary>
    /// <returns>An object that represents the heading of the sheet.</returns>
    public ExcelHeading ReadHeading()
    {
        if (!HasHeading)
        {
            throw new ExcelMappingException($"Sheet \"{Name}\" has no heading.");
        }

        if (Heading != null)
        {
            throw new ExcelMappingException($"Already read heading in sheet \"{Name}\".");
        }

        // Read up to the heading row.
        _cursor.MoveTo(_cursor.GetRowIndex(_dataRange.Rows.Start));

        // Create the heading.
        var heading = new ExcelHeading(Reader, DataRange, Configuration);
        Heading = heading;
        return heading;
    }

    /// <summary>
    /// Maps each row of the sheet to an object using a registered mapping. If no map is registered for this
    /// type then the type will be automapped. This method will read the sheet's heading if the sheet has
    /// a heading and the heading has not yet been read.
    /// </summary>
    /// <remarks>
    /// Reading is forward-only: rows are read from the current position of the sheet. Rows that have
    /// already been read (for example by a previous call to this method or <see cref="ReadRow{T}"/>)
    /// are not returned again, so calling this method after all rows have been read returns an empty
    /// sequence. Use <see cref="ReadRows{T}(int, int)"/> to re-read a specific range of rows.
    /// </remarks>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <returns>A list of objects of type T mapped from each row in the sheet.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
    {
        Worksheet.ThrowIfClosed();

        // Read the heading if we haven't already - this validates the sheet structure
        if (HasHeading && Heading == null)
        {
            ReadHeading();
        }

        // Now delegate to iterator method for lazy evaluation
        return ReadRowsIterator<T>();
    }

    /// <summary>
    /// Iterator method for reading rows. Separated from ReadRows to enable eager validation.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    private IEnumerable<T> ReadRowsIterator<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
    {
        while (TryReadRow(out T? row))
        {
            yield return row;
        }
    }

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
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int startIndex, int count)
    {
        Worksheet.ThrowIfClosed();
        ThrowHelpers.ThrowIfNegative(startIndex, nameof(startIndex));
        if (HasHeading)
        {
            ThrowHelpers.ThrowIfLessThanOrEqual(startIndex, HeadingIndex, nameof(startIndex));
        }
        ThrowHelpers.ThrowIfNegative(count, nameof(count));

        // Read the heading if we haven't already - this validates the sheet structure
        if (HasHeading && Heading == null)
        {
            ReadHeading();
        }

        // The underlying reader is forward-only. If we have already read past the start index,
        // rewind so that the requested rows are returned rather than whatever follows.
        if (startIndex - 1 < CurrentRowIndex)
        {
            _cursor.Reset();
        }

        // Skip to the start index
        _cursor.MoveTo(startIndex - 1);

        // Handle zero count case.
        if (count == 0)
        {
            return [];
        }

        // Now delegate to iterator method for lazy evaluation
        return ReadRowsRangeIterator<T>(startIndex, count);
    }

    /// <summary>
    /// Iterator method for reading a range of rows. Separated from ReadRows to enable eager validation.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    private IEnumerable<T> ReadRowsRangeIterator<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int startIndex, int count)
    {
        for (int i = 0; i < count; i++)
        {
            if (!TryReadRow(out T? row))
            {
                throw new ExcelMappingException($"Sheet \"{Name}\" does not have row {startIndex + i}.");
            }

            yield return row;
        }
    }

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
    public T ReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
    {
        if (!TryReadRow<T>(out var value))
        {
            throw new ExcelMappingException($"No more rows in sheet \"{Name}\".");
        }

        return value;
    }

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
    public bool TryReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>([NotNullWhen(true)] out T? value)
    {
        if (!AdvanceToNextRow())
        {
            value = default;
            return false;
        }

        return MapCurrentRow(out value);
    }

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
    public IEnumerable<ExcelRowResult<T>> ReadRowResults<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
    {
        Worksheet.ThrowIfClosed();

        // Read the heading now, as ReadRows does, so that problems with it are reported immediately.
        if (HasHeading && Heading == null)
        {
            ReadHeading();
        }

        return ReadRowResultsIterator<T>();
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    private IEnumerable<ExcelRowResult<T>> ReadRowResultsIterator<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
    {
        // Reuse one list for every row: valid rows (usually most of them) then don't allocate, and
        // only rows with errors get their own copy.
        var errors = new List<ExcelMappingException>();
        while (AdvanceToNextRow())
        {
            T? value;
            bool mapped;
            CollectedErrors = errors;
            try
            {
                mapped = MapCurrentRow(out value);
            }
            finally
            {
                CollectedErrors = null;
            }

            if (errors.Count == 0)
            {
                yield return new ExcelRowResult<T>(CurrentRowIndex, mapped ? value : default, null);
            }
            else
            {
                var rowErrors = errors.ToArray();
                errors.Clear();
                yield return new ExcelRowResult<T>(CurrentRowIndex, default, rowErrors);
            }
        }
    }

    /// <summary>
    /// When not null, mapping errors for the current row are added to this list instead of being
    /// thrown. Set while reading rows with <see cref="ReadRowResults{T}"/>.
    /// </summary>
    internal List<ExcelMappingException>? CollectedErrors { get; private set; }

    private bool AdvanceToNextRow()
    {
        // Read the heading first, as ReadRows does, so that the heading row isn't mapped as data
        // and columns can be found by name.
        if (HasHeading && Heading == null)
        {
            Worksheet.ThrowIfClosed();

            // A sheet without enough rows for a heading has no rows to read either.
            if (!_cursor.TryMoveTo(_cursor.GetRowIndex(_dataRange.Rows.Start)))
            {
                return false;
            }

            ReadHeading();
        }

        return _cursor.MoveNext(_dataRange, FirstDataRow);
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    private bool MapCurrentRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>([NotNullWhen(true)] out T? value)
    {
        var classMap = GetOrCreateClassMap<T>();

        // Map using the configured culture, if any. Mappers without their own format provider use the
        // current culture, so set it for the duration of this row only, not while the caller's code
        // runs between rows. Only switch if it differs, as setting the culture has a small cost.
        var culture = Configuration.Culture;
        var previousCulture = CultureInfo.CurrentCulture;
        var switchCulture = culture is not null && !ReferenceEquals(culture, previousCulture);
        bool result;
        object? valueObject;
        if (switchCulture)
        {
            CultureInfo.CurrentCulture = culture!;
            try
            {
                result = classMap.TryGetValue(this, CurrentRowIndex, Reader, null, out valueObject);
            }
            finally
            {
                CultureInfo.CurrentCulture = previousCulture;
            }
        }
        else
        {
            result = classMap.TryGetValue(this, CurrentRowIndex, Reader, null, out valueObject);
        }

        // If we've been asked to validate with data annotations, do so now.
        if (Configuration.ValidateDataAnnotations && result && valueObject is not null)
        {
            if (CollectedErrors is { } errors)
            {
                // Report every validation failure rather than only the first.
                var validationResults = new List<ValidationResult>();
                if (!Validator.TryValidateObject(valueObject, new ValidationContext(valueObject), validationResults, validateAllProperties: true))
                {
                    foreach (var validationResult in validationResults)
                    {
                        errors.Add(new ExcelMappingException(validationResult.ErrorMessage ?? "The value is invalid.", this, CurrentRowIndex, -1, new ValidationException(validationResult, null, valueObject)));
                    }
                }
            }
            else
            {
                Validator.ValidateObject(valueObject, new ValidationContext(valueObject), validateAllProperties: true);
            }
        }

        // valueObject is null if the row couldn't be mapped while collecting errors.
        value = valueObject is null ? default : (T)valueObject;
        return result;
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    internal ExcelClassMap GetOrCreateClassMap<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
    {
        if (!Configuration.TryGetClassMap(typeof(T), out var classMap))
        {
            // Sheets without a heading can still be auto-mapped if the map doesn't look up columns by
            // name, for example primitives (read from the first column) or members with ExcelColumnIndex.
            // Otherwise reading a named column throws a clear error.
            if (!AutoMapper.TryCreateClassMap<T>(EmptyValueStrategy.Error, out var map))
            {
                throw new ExcelMappingException($"Cannot auto-map type \"{typeof(T)}\".");
            }

            classMap = Configuration.RegisterAutoMappedClassMap(typeof(T), map);
        }

        RootMap = classMap;

        return classMap;
    }

    /// <summary>
    /// The map that rows of this sheet are being mapped with (the class map of the type passed to
    /// <c>ReadRow</c>, <c>ReadRows</c> or <c>ValidateHeading</c>), as opposed to the nested maps
    /// of its members. Extra columns exclude every column it reads (see
    /// <see cref="ExtraColumnsReaderFactory"/>).
    /// </summary>
    /// <remarks>
    /// Members bound to the sheet keep their extra columns reader, which is correct as a nested
    /// class map belongs to one parent: <c>WithClassMap</c> copies a class map's members rather
    /// than sharing the map. The reader factory can be shared, so its readers are kept per root map.
    /// </remarks>
    internal IMap? RootMap { get; private set; }

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
    public void ValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
    {
        if (!TryValidateHeading<T>(out var problems))
        {
            throw new ExcelMappingException(
                $"Sheet \"{Name}\" is missing columns needed to map type \"{typeof(T)}\": {string.Join("; ", problems)}.");
        }
    }

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
    public bool TryValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(out IReadOnlyList<string> problems)
    {
        Worksheet.ThrowIfClosed();

        if (HasHeading && Heading == null)
        {
            ReadHeading();
        }

        var foundProblems = new List<string>();
        MapValidation.FindMissingColumns(GetOrCreateClassMap<T>(), this, null, foundProblems);
        problems = foundProblems;
        return foundProblems.Count == 0;
    }

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
    public ExcelHeading FindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => FindHeading<T>(HeadingSearch.DefaultMaxRows);

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
    public ExcelHeading FindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int maxRows)
    {
        if (!TryFindHeading<T>(maxRows, out var heading, out var error))
        {
            throw new ExcelMappingException(error);
        }

        return heading;
    }

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
        => TryFindHeading<T>(HeadingSearch.DefaultMaxRows, out heading, out _);

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
        => TryFindHeading<T>(maxRows, out heading, out _);

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    internal bool TryFindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int maxRows, [NotNullWhen(true)] out ExcelHeading? heading, [NotNullWhen(false)] out string? error)
    {
        Worksheet.ThrowIfClosed();
        ThrowHelpers.ThrowIfNegativeOrZero(maxRows, nameof(maxRows));
        if (Heading != null)
        {
            throw new InvalidOperationException("The heading has already been read.");
        }

        return HeadingSearch.TryFind<T>(this, _cursor, maxRows, out heading, out error);
    }

    /// <summary>
    /// Uses a row found by <see cref="HeadingSearch"/> as the heading: the data starts there.
    /// </summary>
    internal void UseHeading(int rowIndex, ExcelHeading heading)
    {
        _hasHeading = true;
        _dataRange = new ExcelRange(rowIndex.._dataRange.Rows.End, _dataRange.Columns);
        Heading = heading;
    }

    /// <summary>
    /// The first row that can be data, if the heading is above the rows to read (see
    /// <see cref="ExcelRegion.WithHeadingAt(int)"/>). Rows between the heading and this row are skipped.
    /// </summary>
    internal Index? FirstDataRow { get; set; }

    /// <summary>
    /// The position in the worksheet of column 0, for a region (see <see cref="ExcelRegion"/>):
    /// columns chosen by position count from the region's first column.
    /// </summary>
    internal int ColumnOffset { get; set; }

    /// <summary>
    /// The position in the worksheet just after the last column that can be chosen by position.
    /// </summary>
    internal int ColumnEnd { get; set; } = int.MaxValue;

    /// <summary>
    /// Converts a column position chosen by the user (for example with <see cref="ExcelColumnIndexAttribute"/>)
    /// to its position in the worksheet, or -1 if the sheet (or region) doesn't have that column.
    /// </summary>
    internal int ToSheetColumn(int columnIndex)
    {
        var sheetColumnIndex = (long)ColumnOffset + columnIndex;
        return sheetColumnIndex < Math.Min(ColumnEnd, NumberOfColumns) ? (int)sheetColumnIndex : -1;
    }

    /// <summary>
    /// Gets the reader a factory creates for this sheet, creating it the first time.
    /// </summary>
    internal ICellReader? GetCellReader(ICellReaderFactory factory) => _bindings.GetCellReader(factory);

    /// <summary>
    /// Gets the cells reader a factory creates for this sheet, creating it the first time.
    /// </summary>
    internal ICellsReader? GetCellsReader(ICellsReaderFactory factory) => _bindings.GetCellsReader(factory);

    /// <summary>
    /// Gets the members of a class map bound to this sheet.
    /// </summary>
    internal BoundMember[] GetBoundMembers(ExcelClassMap classMap) => _bindings.GetBoundMembers(classMap);
}
