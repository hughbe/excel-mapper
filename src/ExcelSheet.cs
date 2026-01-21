using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
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
        if (reader.VisibleState == "visible")
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
        ReadUpTo(HeadingIndex);

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
    /// a heading and the heading has not yet been read.
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
    /// type then the type will be automapped. This method will not read the sheet's heading if the sheet has a
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
    /// type then the type will be automapped. This method will not read the sheet's heading if the sheet has a
    /// heading and the heading has not yet been read.
    /// </summary>
    /// <typeparam name="T">The type of the object to map a single row to.</typeparam>
    /// <param name="value">An object of type T mapped from a single row in the sheet.</param>
    /// <returns>False if there are no more rows in the sheet or the row cannot be mapped to an object, else false.</returns>
    public bool TryReadRow<T>([NotNullWhen(true)] out T? value)
    {
        if (!MoveToNextRow())
        {
            value = default;
            return false;
        }

        if (!Importer.Configuration.TryGetClassMap<T>(out var classMap))
        {
            if (!HasHeading)
            {
                throw new ExcelMappingException($"Cannot auto-map type \"{typeof(T)}\" as the sheet has no heading.");
            }

            if (!AutoMapper.TryCreateClassMap<T>(FallbackStrategy.ThrowIfPrimitive, out var map))
            {
                throw new ExcelMappingException($"Cannot auto-map type \"{typeof(T)}\".");
            }

            classMap = map;
            Importer.Configuration.RegisterClassMap(typeof(T), classMap);
        }

        var result = classMap.TryGetValue(this, CurrentRowIndex, Reader, null, out var valueObject);
        // If we've been asked to validate with data annotations, do so now.
        if (Importer.Configuration.ValidateDataAnnotations && result && valueObject is not null)
        {
            Validator.ValidateObject(valueObject, new ValidationContext(valueObject), validateAllProperties: true);
        }

        value = (T?)valueObject;
        return result;
    }

    private bool MoveToNextRow()
    {
        if (Reader.IsClosed)
        {
            throw new ExcelMappingException($"The underlying reader is closed.");
        }

        // Ensure we're on the correct sheet and row.
        ReadUpTo(CurrentRowIndex);

        // Check if we've reached the end of the data range.
        if (CurrentRowIndex + 1 >= _dataRange.Rows.End.GetOffset(Reader.RowCount))
        {
            return false;
        }

        // Read the next row.
        Reader.Read();
        CurrentRowIndex++;

        if (Importer.Configuration.SkipBlankLines)
        {
            bool RowEmpty()
            {
                for (int i = 0; i < Reader.FieldCount; i++)
                {
                    var value = Reader.GetValue(i);
                    if (value is not null && !(value is string stringValue && string.IsNullOrEmpty(stringValue)))
                    {
                        return false;
                    }
                }

                return true;
            }

            while (RowEmpty())
            {
                if (!Reader.Read())
                {
                    return false;
                }

                CurrentRowIndex++;
            }
        }

        return true;
    }

    internal void ReadUpTo(int index)
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
        }

        // Read up to the correct row.
        for (var i = CurrentRowIndex; i < index; i++)
        {
            if (Reader.IsClosed)
            {
                throw new ExcelMappingException($"The underlying reader is closed.");
            }
            if (!Reader.Read())
            {
                throw new ExcelMappingException($"Sheet \"{Name}\" does not have row {index}.");
            }

            CurrentRowIndex++;
        }
    }
}
