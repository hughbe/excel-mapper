using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ExcelDataReader;
using ExcelMapper.Readers;

namespace ExcelMapper.Reading;

/// <summary>
/// What a sheet or a region reads: a worksheet, its data range and heading, the row it's on and
/// what's been worked out about its columns. An <see cref="ExcelSheet"/> reads through a view of
/// its own, and a region (<see cref="ExcelRegion"/>) through a view with the region's range and
/// heading. Internal code gets the view, which the user's <see cref="ExcelSheet"/> is
/// <see cref="Sheet"/> of.
/// </summary>
internal sealed class SheetView
{
    public SheetView(Worksheet worksheet, ExcelSheet sheet)
    {
        Worksheet = worksheet;
        Sheet = sheet;
        _cursor = new RowCursor(worksheet);
        _bindings = new SheetBindings(this);
    }

    /// <summary>
    /// Gets what's fixed about the worksheet: its name, index and reader.
    /// </summary>
    public Worksheet Worksheet { get; }

    /// <summary>
    /// Gets the user's sheet: the sheet itself, or the sheet a region is of. Extension points and
    /// exceptions are given this sheet.
    /// </summary>
    public ExcelSheet Sheet { get; }

    private readonly RowCursor _cursor;
    private readonly SheetBindings _bindings;

    /// <inheritdoc cref="ExcelSheet.Name"/>
    public string Name => Worksheet.Name;



    /// <inheritdoc cref="ExcelSheet.NumberOfColumns"/>
    public int NumberOfColumns => Worksheet.NumberOfColumns;

    private bool _hasHeading = true;

    /// <inheritdoc cref="ExcelSheet.HasHeading"/>
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

    /// <inheritdoc cref="ExcelSheet.HeadingIndex"/>
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




    /// <inheritdoc cref="ExcelSheet.DataRange"/>
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

    /// <inheritdoc cref="ExcelSheet.Heading"/>
    public ExcelHeading? Heading { get; set; }

    /// <inheritdoc cref="ExcelSheet.CurrentRowIndex"/>
    public int CurrentRowIndex => _cursor.RowIndex;

    /// <summary>
    /// The importer's culture, if set (see <see cref="ExcelImporterConfiguration.Culture"/>).
    /// </summary>
    public CultureInfo? Culture => Settings.Culture;

    private ReadSettings? _settings;

    /// <summary>
    /// The importer's settings, taken when this sheet first needs them: when it starts reading.
    /// </summary>
    public ReadSettings Settings => _settings ??= new ReadSettings(Worksheet.Configuration);

    public ExcelImporterConfiguration Configuration => Worksheet.Configuration;

    private IExcelDataReader Reader => Worksheet.Reader;

    /// <inheritdoc cref="ExcelSheet.ReadHeading"/>
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
        var heading = new ExcelHeading(Reader, DataRange, Settings.MaxColumnsPerSheet);
        Heading = heading;
        return heading;
    }

    /// <inheritdoc cref="ExcelSheet.ReadRows{T}()"/>
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

    /// <inheritdoc cref="ExcelSheet.ReadRows{T}(int, int)"/>
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

    /// <inheritdoc cref="ExcelSheet.ReadRow{T}"/>
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

    /// <inheritdoc cref="ExcelSheet.TryReadRow{T}(out T)"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>([NotNullWhen(true)] out T? value)
    {
        if (!AdvanceToNextRow())
        {
            value = default;
            return false;
        }

        return MapCurrentRow(null, out value);
    }

    /// <inheritdoc cref="ExcelSheet.ReadRowResults{T}"/>
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
            var mapped = MapCurrentRow(errors, out T? value);

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

        return _cursor.MoveNext(_dataRange, FirstDataRow, Settings.SkipBlankLines);
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    private bool MapCurrentRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(List<ExcelMappingException>? errors, [NotNullWhen(true)] out T? value)
    {
        var classMap = GetOrCreateClassMap<T>();

        // The importer's culture, if set, is given to each cell (see ValuePipeline<T>.Map), not set
        // as the thread's culture.
        var result = classMap.TryGetValue(this, CurrentRowIndex, Reader, null, errors, out var valueObject);

        // If we've been asked to validate with data annotations, do so now.
        if (Settings.ValidateDataAnnotations && result && valueObject is not null)
        {
            if (errors is not null)
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
    public ExcelClassMap GetOrCreateClassMap<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
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

        // The class map can't change once it's used with a sheet, so what's worked out from it
        // (its plans and bindings) stays right.
        classMap.Freeze();
        RootMap = classMap;

        return classMap;
    }

    /// <summary>
    /// The map that rows of this sheet are being mapped with (the class map of the type passed to
    /// <c>ReadRow</c>, <c>ReadRows</c> or <c>ValidateHeading</c>), as opposed to the nested maps
    /// of its members. Extra columns exclude every column it reads (see
    /// <see cref="ExtraColumns"/>).
    /// </summary>
    /// <remarks>
    /// Members bound to the sheet keep their extra columns reader, which is correct as a nested
    /// class map belongs to one parent: <c>WithClassMap</c> copies a class map's members rather
    /// than sharing the map. The extra columns can be shared, so their readers are kept per root map.
    /// </remarks>
    public IMap? RootMap { get; private set; }

    /// <inheritdoc cref="ExcelSheet.ValidateHeading{T}"/>
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

    /// <inheritdoc cref="ExcelSheet.TryValidateHeading{T}(out IReadOnlyList{string})"/>
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

    /// <inheritdoc cref="ExcelSheet.FindHeading{T}()"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public ExcelHeading FindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => FindHeading<T>(HeadingSearch.DefaultMaxRows);

    /// <inheritdoc cref="ExcelSheet.FindHeading{T}(int)"/>
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

    /// <inheritdoc cref="ExcelSheet.TryFindHeading{T}(out ExcelHeading)"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryFindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>([NotNullWhen(true)] out ExcelHeading? heading)
        => TryFindHeading<T>(HeadingSearch.DefaultMaxRows, out heading, out _);

    /// <inheritdoc cref="ExcelSheet.TryFindHeading{T}(int, out ExcelHeading)"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryFindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int maxRows, [NotNullWhen(true)] out ExcelHeading? heading)
        => TryFindHeading<T>(maxRows, out heading, out _);

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryFindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int maxRows, [NotNullWhen(true)] out ExcelHeading? heading, [NotNullWhen(false)] out string? error)
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
    public void UseHeading(int rowIndex, ExcelHeading heading)
    {
        _hasHeading = true;
        _dataRange = new ExcelRange(rowIndex.._dataRange.Rows.End, _dataRange.Columns);
        Heading = heading;
    }

    /// <summary>
    /// The first row that can be data, if the heading is above the rows to read (see
    /// <see cref="ExcelRegion.WithHeadingAt(int)"/>). Rows between the heading and this row are skipped.
    /// </summary>
    public Index? FirstDataRow { get; set; }

    /// <summary>
    /// The position in the worksheet of column 0, for a region (see <see cref="ExcelRegion"/>):
    /// columns chosen by position count from the region's first column.
    /// </summary>
    public int ColumnOffset { get; set; }

    /// <summary>
    /// The position in the worksheet just after the last column that can be chosen by position.
    /// </summary>
    public int ColumnEnd { get; set; } = int.MaxValue;

    /// <summary>
    /// Converts a column position chosen by the user (for example with <see cref="ExcelColumnIndexAttribute"/>)
    /// to its position in the worksheet, or -1 if the sheet (or region) doesn't have that column.
    /// </summary>
    public int ToSheetColumn(int columnIndex)
    {
        var sheetColumnIndex = (long)ColumnOffset + columnIndex;
        return sheetColumnIndex < Math.Min(ColumnEnd, NumberOfColumns) ? (int)sheetColumnIndex : -1;
    }

    /// <summary>
    /// Finds the column a single value is read from, once for this sheet. -1 if the sheet doesn't
    /// have it.
    /// </summary>
    public int FindColumn(Columns columns) => _bindings.FindColumn(columns);

    /// <summary>
    /// Gets the reader of the cells a collection or dictionary is read from, once for this sheet.
    /// Null if the sheet doesn't have the columns.
    /// </summary>
    public ICellsReader? GetCellsReader(Columns columns) => _bindings.GetCellsReader(columns);

    /// <summary>
    /// Gets the members of a class map bound to this sheet.
    /// </summary>
    public BoundMember[] GetBoundMembers(ExcelClassMap classMap) => _bindings.GetBoundMembers(classMap);
}
