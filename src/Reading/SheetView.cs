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

    /// <summary>
    /// Gets the heading, to find columns by name, or throws if the rows have none.
    /// </summary>
    public ExcelHeading GetHeadingForNames()
        => Heading ?? throw NoHeading("to find columns by name", " Choose the columns by index or letter instead.");

    /// <summary>
    /// Gets the heading, to read every column with a name, or throws if the rows have none.
    /// </summary>
    public ExcelHeading GetHeadingForAllColumns()
        => Heading ?? throw NoHeading("to read every named column from", " Choose the columns by index instead.");

    /// <summary>
    /// Gets the heading, whose names are a dictionary's keys, or throws if the rows have none.
    /// </summary>
    public ExcelHeading GetHeadingForKeys()
        => Heading ?? throw NoHeading("to name a dictionary's keys", " A dictionary's keys are its columns' names, so it can only be read from rows with a heading.");

    /// <summary>
    /// Gets the heading, to find the extra columns no member reads, or throws if the rows have none.
    /// </summary>
    public ExcelHeading GetHeadingForExtraColumns()
        => Heading ?? throw NoHeading("to read extra columns from", " Extra columns are the heading's columns that no other member reads.");

    // The sheet's name follows the message, then the hint.
    private ExcelMappingException NoHeading(string purpose, string hint)
        => new($"There's no heading {purpose}", this, -1, -1, null, hint);

    /// <inheritdoc cref="ExcelSheet.CurrentRowIndex"/>
    public int CurrentRowIndex => _cursor.RowIndex;

    /// <summary>
    /// The importer's culture, if set (see <see cref="ExcelImporterConfiguration.Culture"/>).
    /// </summary>
    public CultureInfo Culture => Settings.Culture;

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
            throw new ExcelMappingException($"{ExcelMappingException.DescribeSheet(Name, capitalize: true)} has no heading.");
        }

        if (Heading != null)
        {
            throw new ExcelMappingException($"The heading of {ExcelMappingException.DescribeSheet(Name)} has already been read.");
        }

        // Read up to the heading row.
        var headingRowIndex = _cursor.GetRowIndex(_dataRange.Rows.Start);
        if (!_cursor.TryMoveTo(headingRowIndex))
        {
            // The cursor is on the sheet's last row, if any.
            var rowCount = _cursor.RowIndex + 1;
            throw new ExcelMappingException(rowCount == 0
                ? $"{ExcelMappingException.DescribeSheet(Name, capitalize: true)} is empty, so it has no heading. Set HasHeading to false to read a sheet without one."
                : $"{ExcelMappingException.DescribeSheet(Name, capitalize: true)} has no row {headingRowIndex + 1} to read the heading from: it has {rowCount} row{(rowCount == 1 ? "" : "s")}.");
        }

        // Create the heading.
        var heading = new ExcelHeading(Reader, DataRange, Settings);
        Heading = heading;
        return heading;
    }

    /// <inheritdoc cref="ExcelSheet.ReadRows{T}()"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => ReadRows(AutoMapped<T>.ClassMap);

    /// <inheritdoc cref="ExcelSheet.ReadRows{T}(ExcelClassMap{T})"/>
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap)
    {
        Worksheet.ThrowIfClosed();

        // Read the heading if we haven't already - this validates the sheet structure
        if (HasHeading && Heading == null)
        {
            ReadHeading();
        }

        // Rows are read lazily, from the row the sheet is on now.
        return ReadRowsIterator(classMap, CurrentRowIndex);
    }

    /// <summary>
    /// Reads the rows after <paramref name="rowIndex"/>. Separated from ReadRows so the heading is
    /// read when ReadRows is called rather than when the rows are enumerated.
    /// </summary>
    private IEnumerable<T> ReadRowsIterator<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, int rowIndex)
    {
        // Each enumeration starts from the same row, so the result can be enumerated more than
        // once, as with any IEnumerable<T>: calling Any() and then ToList() doesn't lose a row.
        _cursor.MoveTo(rowIndex);
        ExcelClassMap<T>? map = null;
        while (AdvanceToNextRow())
        {
            map = UseClassMapForRow(map, classMap);
            if (!MapCurrentRow(map, null, out T? row))
            {
                yield break;
            }

            yield return row;
        }
    }

    /// <inheritdoc cref="ExcelSheet.ReadRows{T}(int, int)"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int startIndex, int count) => ReadRows(AutoMapped<T>.ClassMap, startIndex, count);

    /// <inheritdoc cref="ExcelSheet.ReadRows{T}(ExcelClassMap{T}, int, int)"/>
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, int startIndex, int count)
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
        return ReadRowsRangeIterator(classMap, startIndex, count);
    }

    /// <summary>
    /// Iterator method for reading a range of rows. Separated from ReadRows to enable eager validation.
    /// </summary>
    private IEnumerable<T> ReadRowsRangeIterator<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, int startIndex, int count)
    {
        // Each enumeration reads the same rows.
        _cursor.MoveTo(startIndex - 1);
        ExcelClassMap<T>? map = null;
        for (int i = 0; i < count; i++)
        {
            if (!AdvanceToNextRow() || !MapCurrentRow(map = UseClassMapForRow(map, classMap), null, out T? row))
            {
                throw new ExcelMappingException($"{ExcelMappingException.DescribeSheet(Name, capitalize: true)} does not have row {startIndex + i + 1}.");
            }

            yield return row;
        }
    }

    /// <inheritdoc cref="ExcelSheet.ReadRow{T}()"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public T ReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => ReadRow(AutoMapped<T>.ClassMap);

    /// <inheritdoc cref="ExcelSheet.ReadRow{T}(ExcelClassMap{T})"/>
    public T ReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap)
    {
        if (!TryReadRow(classMap, out T? value))
        {
            throw new ExcelMappingException($"There are no more rows in {ExcelMappingException.DescribeSheet(Name)}.");
        }

        return value;
    }

    /// <inheritdoc cref="ExcelSheet.TryReadRow{T}(out T)"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>([NotNullWhen(true)] out T? value) => TryReadRow(AutoMapped<T>.ClassMap, out value);

    /// <inheritdoc cref="ExcelSheet.TryReadRow{T}(ExcelClassMap{T}, out T)"/>
    public bool TryReadRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, [NotNullWhen(true)] out T? value)
    {
        if (!AdvanceToNextRow())
        {
            value = default;
            return false;
        }

        return MapCurrentRow(classMap(this), null, out value);
    }

    /// <inheritdoc cref="ExcelSheet.ReadRowResults{T}()"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<ExcelRowResult<T>> ReadRowResults<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => ReadRowResults(AutoMapped<T>.ClassMap);

    /// <inheritdoc cref="ExcelSheet.ReadRowResults{T}(ExcelClassMap{T})"/>
    public IEnumerable<ExcelRowResult<T>> ReadRowResults<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap)
    {
        Worksheet.ThrowIfClosed();

        // Read the heading now, as ReadRows does, so that problems with it are reported immediately.
        if (HasHeading && Heading == null)
        {
            ReadHeading();
        }

        // A column the sheet doesn't have is a problem with the sheet rather than its rows, so
        // it's reported once, as ValidateHeading does, rather than as an error of every row.
        ValidateHeading(classMap);

        return ReadRowResultsIterator(classMap, CurrentRowIndex);
    }

    private IEnumerable<ExcelRowResult<T>> ReadRowResultsIterator<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, int rowIndex)
    {
        // Each enumeration starts from the same row, as for ReadRows.
        _cursor.MoveTo(rowIndex);

        // Reuse one list for every row: valid rows (usually most of them) then don't allocate, and
        // only rows with errors get their own copy.
        var target = new ExceptionListTarget();
        var errors = target.Exceptions;
        ExcelClassMap<T>? map = null;
        while (AdvanceToNextRow())
        {
            map = UseClassMapForRow(map, classMap);
            var mapped = MapCurrentRow(map, target, out T? value);

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

    /// <inheritdoc cref="ExcelSheet.ReadRows{T}(IExcelErrorSink)"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(IExcelErrorSink errors) => ReadRows(AutoMapped<T>.ClassMap, errors);

    /// <inheritdoc cref="ExcelSheet.ReadRows{T}(ExcelClassMap{T}, IExcelErrorSink)"/>
    public IEnumerable<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, IExcelErrorSink errors)
    {
        ThrowHelpers.ThrowIfNull(errors, nameof(errors));
        Worksheet.ThrowIfClosed();

        // As ReadRowResults does: problems with the heading are reported once, now.
        if (HasHeading && Heading == null)
        {
            ReadHeading();
        }

        ValidateHeading(classMap);
        return ReadRowsIterator(classMap, new SinkTarget(errors), CurrentRowIndex);
    }

    private IEnumerable<T> ReadRowsIterator<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, SinkTarget target, int rowIndex)
    {
        // Each enumeration starts from the same row, as for ReadRows.
        _cursor.MoveTo(rowIndex);
        ExcelClassMap<T>? map = null;
        while (AdvanceToNextRow())
        {
            map = UseClassMapForRow(map, classMap);

            // A row with errors is reported to the sink, and skipped: partly mapped objects are
            // never returned.
            var errorCount = target.Count;
            if (MapCurrentRow(map, target, out T? value) && target.Count == errorCount)
            {
                yield return value;
            }
        }
    }

    /// <summary>
    /// Gets the class map to map a row with, while reading several rows: looked up (and frozen)
    /// on the first row, as <paramref name="classMap"/> does, and reused for the rest. Looking up a
    /// registered or auto-mapped class map for every row cost about 5% of reading narrow rows.
    /// </summary>
    private ExcelClassMap<T> UseClassMapForRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T>? map, Func<SheetView, ExcelClassMap<T>> classMap)
    {
        if (map is null)
        {
            return classMap(this);
        }

        // Another read of this sheet, between two rows of this one, may have used another map.
        RootMap = map;
        return map;
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

    // Data annotations are validated by reflection over T's properties, which ExcelClassMap<T> keeps.
    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "The public properties of T, which Validator validates, are kept by ExcelClassMap<T>'s annotation.")]
    private bool MapCurrentRow<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap, ErrorTarget? errors, [NotNullWhen(true)] out T? value)
    {
        // The importer's culture, if set, is given to each cell (see CellPipeline<T>.Map), not set
        // as the thread's culture.
        var result = classMap.TryGetValue(this, CurrentRowIndex, Reader, null, errors, out value);

        // If we've been asked to validate with data annotations, do so now.
        if (Settings.ValidateDataAnnotations && result && value is not null)
        {
            object valueObject = value;
            if (errors is not null)
            {
                // Report every validation failure rather than only the first.
                var validationResults = new List<ValidationResult>();
                if (!Validator.TryValidateObject(valueObject, new ValidationContext(valueObject), validationResults, validateAllProperties: true))
                {
                    foreach (var validationResult in validationResults)
                    {
                        errors.Add(new ExcelMappingException(validationResult.ErrorMessage ?? "The value is invalid.", this, CurrentRowIndex, -1, new ValidationException(validationResult, null, valueObject)), ExcelMappingErrorKind.DataAnnotation);
                    }
                }
            }
            else
            {
                Validator.ValidateObject(valueObject, new ValidationContext(valueObject), validateAllProperties: true);
            }
        }

        return result;
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public ExcelClassMap<T> GetOrCreateClassMap<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
    {
        if (!Configuration.TryGetClassMap(typeof(T), out var classMap))
        {
            // Sheets without a heading can still be auto-mapped if the map doesn't look up columns by
            // name, for example primitives (read from the first column) or members with ExcelColumnIndex.
            // Otherwise reading a named column throws a clear error.
            if (!AutoMapper.TryCreateClassMap<T>(EmptyValueStrategy.Error, out var map))
            {
                throw new ExcelMappingException($"Cannot auto-map type {TypeNames.Get(typeof(T))}.");
            }

            classMap = Configuration.RegisterAutoMappedClassMap(typeof(T), map);
        }

        // Class maps are registered for their own type (see ExcelImporterConfiguration.RegisterClassMap).
        return UseClassMap((ExcelClassMap<T>)classMap);
    }

    /// <summary>
    /// Maps rows with <paramref name="classMap"/>: a class map given to a reading method, or the
    /// one registered or auto-mapped for the type.
    /// </summary>
    public ExcelClassMap<T> UseClassMap<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap)
    {
        // The class map can't change once it's used with a sheet, so what's worked out from it
        // (its plans and bindings) stays right.
        classMap.Freeze();
        RootMap = classMap;
        return classMap;
    }

    /// <summary>
    /// Gets the class map a reading method given <paramref name="classMap"/> maps rows with.
    /// </summary>
    public static Func<SheetView, ExcelClassMap<T>> Given<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap)
    {
        ThrowHelpers.ThrowIfNull(classMap, nameof(classMap));
        return sheet => sheet.UseClassMap(classMap);
    }

    /// <summary>
    /// Gets the class map registered for T, or auto-mapped (see <see cref="GetOrCreateClassMap{T}"/>),
    /// for each row: what reading methods without a class map use.
    /// </summary>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public static class AutoMapped<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>
    {
        public static readonly Func<SheetView, ExcelClassMap<T>> ClassMap = static sheet => sheet.GetOrCreateClassMap<T>();
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

    /// <inheritdoc cref="ExcelSheet.ValidateHeading{T}()"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public void ValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>() => ValidateHeading(AutoMapped<T>.ClassMap);

    /// <inheritdoc cref="ExcelSheet.ValidateHeading{T}(ExcelClassMap{T})"/>
    public void ValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap)
    {
        if (!TryValidateHeading(classMap, out var problems))
        {
            throw new ExcelMappingException(
                $"{ExcelMappingException.DescribeSheet(Name, capitalize: true)} is missing columns needed to map type {TypeNames.Get(typeof(T))}: {string.Join("; ", problems)}.");
        }
    }

    /// <inheritdoc cref="ExcelSheet.TryValidateHeading{T}(out IReadOnlyList{string})"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(out IReadOnlyList<string> problems) => TryValidateHeading(AutoMapped<T>.ClassMap, out problems);

    /// <inheritdoc cref="ExcelSheet.TryValidateHeading{T}(ExcelClassMap{T}, out IReadOnlyList{string})"/>
    public bool TryValidateHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, out IReadOnlyList<string> problems)
    {
        Worksheet.ThrowIfClosed();

        if (HasHeading && Heading == null)
        {
            ReadHeading();
        }

        var foundProblems = new List<string>();
        MapValidation.FindMissingColumns(classMap(this), this, null, foundProblems);
        problems = foundProblems;
        return foundProblems.Count == 0;
    }

    /// <inheritdoc cref="ExcelSheet.FindHeading{T}(int)"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public ExcelHeading FindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int maxRows) => FindHeading(AutoMapped<T>.ClassMap, maxRows);

    /// <inheritdoc cref="ExcelSheet.FindHeading{T}(ExcelClassMap{T}, int)"/>
    public ExcelHeading FindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, int maxRows)
    {
        if (!TryFindHeading(classMap, maxRows, out var heading, out var error))
        {
            throw new ExcelMappingException(error);
        }

        return heading;
    }

    /// <inheritdoc cref="ExcelSheet.TryFindHeading{T}(int, out ExcelHeading)"/>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public bool TryFindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(int maxRows, [NotNullWhen(true)] out ExcelHeading? heading, [NotNullWhen(false)] out string? error)
        => TryFindHeading(AutoMapped<T>.ClassMap, maxRows, out heading, out error);

    /// <inheritdoc cref="ExcelSheet.TryFindHeading{T}(ExcelClassMap{T}, int, out ExcelHeading)"/>
    public bool TryFindHeading<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Func<SheetView, ExcelClassMap<T>> classMap, int maxRows, [NotNullWhen(true)] out ExcelHeading? heading, [NotNullWhen(false)] out string? error)
    {
        Worksheet.ThrowIfClosed();
        ThrowHelpers.ThrowIfNegativeOrZero(maxRows, nameof(maxRows));
        if (Heading != null)
        {
            throw new InvalidOperationException("The heading has already been read.");
        }

        return HeadingSearch.TryFind(this, _cursor, classMap(this), maxRows, out heading, out error);
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
    public BoundMember<T>[] GetBoundMembers<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelClassMap<T> classMap) => _bindings.GetBoundMembers(classMap);
}
