using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using ExcelDataReader;

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
    internal ExcelSheet(IExcelDataReader reader, int index, ExcelImporter importer)
    {
        Reader = reader;
        Name = reader.Name;
        // CSV readers have no concept of sheet visibility and return a null state.
        if (reader.VisibleState is null || reader.VisibleState == "visible")
        {
            Visibility = ExcelSheetVisibility.Visible;
        }
        else if (reader.VisibleState == "hidden")
        {
            Visibility = ExcelSheetVisibility.Hidden;
        }
        else
        {
            Visibility = ExcelSheetVisibility.VeryHidden;
        }
        Index = index;
        NumberOfColumns = reader.FieldCount;
        Importer = importer;
    }

    /// <summary>
    /// Gets the name of the sheet.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Gets the visibility of the sheet.
    /// </summary>
    public ExcelSheetVisibility Visibility { get; }

    /// <summary>
    /// Gets the zero-based index of the sheet where 0 is the first sheet in the document.
    /// </summary>
    public int Index { get; }

    /// <summary>
    /// Gets the number of columns in the sheet.
    /// </summary>
    public int NumberOfColumns { get; }

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
    public ExcelHeading? Heading { get; private set; }

    /// <summary>
    /// Gets the index of the row currently being mapped.
    /// </summary>
    public int CurrentRowIndex { get; private set; } = -1;

    private ExcelImporter Importer { get; }

    internal ExcelImporterConfiguration Configuration => Importer.Configuration;

    private IExcelDataReader Reader { get; }

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
        ReadUpTo(GetHeadingRowIndex());

        // Create the heading.
        var heading = new ExcelHeading(Reader, DataRange, Importer.Configuration);
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
    public IEnumerable<T> ReadRows<T>()
    {
        if (Reader.IsClosed)
        {
            throw new ExcelMappingException($"The underlying reader is closed.");
        }

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
    private IEnumerable<T> ReadRowsIterator<T>()
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
    public IEnumerable<T> ReadRows<T>(int startIndex, int count)
    {
        if (Reader.IsClosed)
        {
            throw new ExcelMappingException($"The underlying reader is closed.");
        }
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
            Importer.SheetIndex = -1;
            CurrentRowIndex = -1;
        }

        // Skip to the start index
        ReadUpTo(startIndex - 1);

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
    private IEnumerable<T> ReadRowsRangeIterator<T>(int startIndex, int count)
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
    public T ReadRow<T>()
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
    public bool TryReadRow<T>([NotNullWhen(true)] out T? value)
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
    public IEnumerable<ExcelRowResult<T>> ReadRowResults<T>()
    {
        if (Reader.IsClosed)
        {
            throw new ExcelMappingException($"The underlying reader is closed.");
        }

        // Read the heading now, as ReadRows does, so that problems with it are reported immediately.
        if (HasHeading && Heading == null)
        {
            ReadHeading();
        }

        return ReadRowResultsIterator<T>();
    }

    private IEnumerable<ExcelRowResult<T>> ReadRowResultsIterator<T>()
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
            if (Reader.IsClosed)
            {
                throw new ExcelMappingException($"The underlying reader is closed.");
            }

            // A sheet without enough rows for a heading has no rows to read either.
            if (!TryReadUpTo(GetHeadingRowIndex()))
            {
                return false;
            }

            ReadHeading();
        }

        return MoveToNextRow();
    }

    private bool MapCurrentRow<T>([NotNullWhen(true)] out T? value)
    {
        var classMap = GetOrCreateClassMap<T>();

        // Map using the configured culture, if any. Mappers without their own format provider use the
        // current culture, so set it for the duration of this row only, not while the caller's code
        // runs between rows. Only switch if it differs, as setting the culture has a small cost.
        var culture = Importer.Configuration.Culture;
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
        if (Importer.Configuration.ValidateDataAnnotations && result && valueObject is not null)
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

    private IMap GetOrCreateClassMap<T>()
    {
        if (!Importer.Configuration.TryGetClassMap<T>(out var classMap))
        {
            // Sheets without a heading can still be auto-mapped if the map doesn't look up columns by
            // name, for example primitives (read from the first column) or members with ExcelColumnIndex.
            // Otherwise reading a named column throws a clear error.
            if (!AutoMapper.TryCreateClassMap<T>(FallbackStrategy.ThrowIfPrimitive, out var map))
            {
                throw new ExcelMappingException($"Cannot auto-map type \"{typeof(T)}\".");
            }

            classMap = Importer.Configuration.RegisterAutoMappedClassMap(typeof(T), map);
        }

        return classMap;
    }

    /// <summary>
    /// Checks that the sheet has every column needed to map rows to <typeparamref name="T"/>, and
    /// throws an <see cref="ExcelMappingException"/> listing all of the missing columns if not.
    /// Call this before reading rows to report every missing column at once, rather than one at a
    /// time when the first row is read. Optional members are not required. This reads the heading
    /// if it has not been read yet.
    /// </summary>
    /// <typeparam name="T">The type that rows will be mapped to.</typeparam>
    public void ValidateHeading<T>()
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
    public bool TryValidateHeading<T>(out IReadOnlyList<string> problems)
    {
        if (Reader.IsClosed)
        {
            throw new ExcelMappingException($"The underlying reader is closed.");
        }

        if (HasHeading && Heading == null)
        {
            ReadHeading();
        }

        var foundProblems = new List<string>();
        MapValidation.FindMissingColumns(GetOrCreateClassMap<T>(), this, null, foundProblems);
        problems = foundProblems;
        return foundProblems.Count == 0;
    }

    private bool MoveToNextRow()
    {
        if (Reader.IsClosed)
        {
            throw new ExcelMappingException($"The underlying reader is closed.");
        }

        // Ensure we're on the correct sheet and row.
        ReadUpTo(CurrentRowIndex);

        // Check if we've reached the end of the data range. Only count the rows for ranges
        // relative to the end of the sheet: otherwise read until there are no more rows.
        var rows = _dataRange.Rows;
        var endRowIndex = rows.End.IsFromEnd && rows.End.Value == 0 ? int.MaxValue : rows.End.GetOffset(GetRowCount());
        if (CurrentRowIndex + 1 >= endRowIndex)
        {
            return false;
        }

        // Skip any rows before the start of the data range. If the sheet has a heading, the
        // heading row is the start of the data range and has already been read.
        var startRowIndex = rows.Start.IsFromEnd ? rows.Start.GetOffset(GetRowCount()) : rows.Start.Value;
        if (CurrentRowIndex + 1 < startRowIndex)
        {
            if (startRowIndex >= endRowIndex || !TryReadUpTo(startRowIndex - 1))
            {
                return false;
            }
        }

        // Read the next row.
        if (!ReadNextRow())
        {
            return false;
        }

        if (Importer.Configuration.SkipBlankLines)
        {
            // Only consider the columns in the data range when checking for blank rows.
            var fieldCount = Reader.FieldCount;
            var startColumnIndex = Math.Min(_dataRange.Columns.Start.GetOffset(fieldCount), fieldCount);
            var endColumnIndex = Math.Min(_dataRange.Columns.End.GetOffset(fieldCount), fieldCount);

            while (IsBlankRow(startColumnIndex, endColumnIndex))
            {
                // Don't skip past the end of the data range.
                if (CurrentRowIndex + 1 >= endRowIndex || !ReadNextRow())
                {
                    return false;
                }
            }
        }

        return true;
    }

    private int GetHeadingRowIndex()
    {
        var start = _dataRange.Rows.Start;
        return start.IsFromEnd ? start.GetOffset(GetRowCount()) : start.Value;
    }

    private int _rowCount = -1;

    /// <summary>
    /// Gets the number of rows in the sheet, ignoring empty lines at the end of a CSV file.
    /// </summary>
    private int GetRowCount()
    {
        // The CSV reader's row count leaves out empty lines, so count the rows by reading the rest
        // of the sheet, once, and then go back to the current row.
        if (!Importer.RowCountExcludesEmptyLines)
        {
            return Reader.RowCount;
        }

        if (_rowCount == -1)
        {
            TryReadUpTo(CurrentRowIndex);
            var currentRowIndex = CurrentRowIndex;
            var rowCount = currentRowIndex + 1;
            while (ReadNextRow())
            {
                rowCount = CurrentRowIndex + 1;
            }

            Importer.SheetIndex = -1;
            TryReadUpTo(currentRowIndex);
            _rowCount = rowCount;
        }

        return _rowCount;
    }

    /// <summary>
    /// The number of rows read from the sheet so far that have a value in any column.
    /// </summary>
    private int _rowsWithValues;

    /// <summary>
    /// Reads the next row of the sheet. Returns false at the end of the sheet, including when
    /// only empty lines remain in a CSV file.
    /// </summary>
    private bool ReadNextRow()
    {
        if (!Reader.Read())
        {
            return false;
        }

        CurrentRowIndex++;
        if (!IsBlankRow(0, Reader.FieldCount))
        {
            _rowsWithValues++;
            return true;
        }

        // The CSV reader's row count leaves out empty lines, so it can be less than the number of
        // rows read. Every row with a value is counted, so once that many rows with values have
        // been read, only empty lines remain: treat them as the end of the sheet, as before. For
        // other readers the row count includes every row, so this is never true.
        return _rowsWithValues < Reader.RowCount;
    }

    private bool IsBlankRow(int startColumnIndex, int endColumnIndex)
    {
        for (int i = startColumnIndex; i < endColumnIndex; i++)
        {
            var value = Reader.GetValue(i);
            if (value is not null && !(value is string stringValue && string.IsNullOrEmpty(stringValue)))
            {
                return false;
            }
        }

        return true;
    }

    internal void ReadUpTo(int index)
    {
        if (!TryReadUpTo(index))
        {
            throw new ExcelMappingException($"Sheet \"{Name}\" does not have row {index}.");
        }
    }

    private bool TryReadUpTo(int index)
    {
        // If we're already on the correct sheet, no need to do anything.
        if (Importer.SheetIndex != Index)
        {
            // Read up to the correct sheet.
            Reader.Reset();
            Importer.SheetIndex = 0;
            
            for (var i = 0; i < Index; i++)
            {
                Reader.NextResult();
                Importer.SheetIndex++;
            }

            CurrentRowIndex = -1;
            _rowsWithValues = 0;
        }

        // Read up to the correct row.
        while (CurrentRowIndex < index)
        {
            if (Reader.IsClosed)
            {
                throw new ExcelMappingException($"The underlying reader is closed.");
            }
            if (!ReadNextRow())
            {
                return false;
            }
        }

        return true;
    }
}
