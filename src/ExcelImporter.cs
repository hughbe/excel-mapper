using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Text;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// An importer that reads the sheets in an Excel file or stream.
/// </summary>
public class ExcelImporter : IDisposable
{
    /// <summary>
    /// Gets the inner reader that the importer wraps. It's internal: the importer's sheets share it,
    /// and only <see cref="RowSource"/> moves it.
    /// </summary>
    internal IExcelDataReader Reader { get; }


    /// <summary>
    /// Gets the configuration for the importer to allow customizing the importer. For example you
    /// can register custom class maps and specify whether sheets have a header.
    /// </summary>
    public ExcelImporterConfiguration Configuration { get; } = new ExcelImporterConfiguration();

    /// <summary>
    /// The stream that was opened by the constructor, if any. This needs to be disposed.
    /// </summary>
    private Stream? _ownedStream = null;

    /// <summary>
    /// Gets the number of sheets in the document.
    /// </summary>
    public int NumberOfSheets => Reader.ResultsCount;

    /// <summary>
    /// The index of the last sheet returned by <see cref="TryReadSheet(out ExcelSheet)"/>, which reads
    /// sheets in order, or -1 to start from the first sheet again.
    /// </summary>
    internal int SheetIndex { get; set; } = -1;

    private RowSource? _rowSource;

    /// <summary>
    /// Owns the reader's position, shared by every sheet of this importer.
    /// </summary>
    internal RowSource RowSource => _rowSource ??= new RowSource(Reader);

    static ExcelImporter()
    {
        // ExcelDataReader needs code page encodings (such as Windows-1252) to read .xls files and
        // CSV files, but .NET Core and later don't include them by default. Registering the
        // provider is global, safe to do more than once, and doesn't change the default encoding.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <summary>
    /// Constructs an importer that reads an Excel file from a path.
    /// </summary>
    /// <param name="path">A path to an Excel file.</param>
    public ExcelImporter(string path) : this(path, ExcelImporterFileType.Excel)
    {
    }

    /// <summary>
    /// Constructs an importer that reads an Excel or csv file from a path.
    /// </summary>
    /// <param name="path">A path to the file to read.</param>
    /// <param name="fileType">The type of file.</param>
    public ExcelImporter(string path, ExcelImporterFileType fileType) : this(path, fileType, null)
    {
    }

    /// <summary>
    /// Constructs an importer that reads an Excel or csv file from a path, with options for the
    /// underlying ExcelDataReader reader such as a password, CSV separators or a fallback encoding.
    /// </summary>
    /// <param name="path">A path to the file to read.</param>
    /// <param name="fileType">The type of file.</param>
    /// <param name="readerConfiguration">Options for the underlying reader, or null for the defaults.</param>
    public ExcelImporter(string path, ExcelImporterFileType fileType, ExcelReaderConfiguration? readerConfiguration)
    {
        ThrowHelpers.ThrowIfNull(path, nameof(path));

        Stream? stream = null;
        try
        {
            stream = File.OpenRead(path);
            Reader = CreateReader(stream, fileType, readerConfiguration);
            _ownedStream = stream;
            stream = null; // Transfer ownership - prevent disposal in catch block
        }
        catch
        {
            stream?.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Constructs an importer that reads an Excel file from a stream.
    /// </summary>
    /// <param name="stream">A stream containing the Excel file bytes.</param>
    public ExcelImporter(Stream stream) : this(stream, ExcelImporterFileType.Excel)
    {
    }

    
    /// <summary>
    /// Constructs an importer that reads an Excel or csv file from a stream.
    /// </summary>
    /// <param name="stream">A stream containing the file to read.</param>
    /// <param name="fileType">The type of file.</param>
    public ExcelImporter(Stream stream, ExcelImporterFileType fileType) : this(stream, fileType, null)
    {
    }

    /// <summary>
    /// Constructs an importer that reads an Excel or csv file from a stream, with options for the
    /// underlying ExcelDataReader reader such as a password, CSV separators or a fallback encoding.
    /// </summary>
    /// <param name="stream">A stream containing the file to read.</param>
    /// <param name="fileType">The type of file.</param>
    /// <param name="readerConfiguration">Options for the underlying reader, or null for the defaults.</param>
    public ExcelImporter(Stream stream, ExcelImporterFileType fileType, ExcelReaderConfiguration? readerConfiguration)
    {
        ThrowHelpers.ThrowIfNull(stream, nameof(stream));

        Reader = CreateReader(stream, fileType, readerConfiguration);
    }

    private static IExcelDataReader CreateReader(Stream stream, ExcelImporterFileType fileType, ExcelReaderConfiguration? readerConfiguration) => fileType switch
    {
        ExcelImporterFileType.Excel => ExcelReaderFactory.CreateReader(stream, readerConfiguration),
        ExcelImporterFileType.Csv => ExcelReaderFactory.CreateCsvReader(stream, readerConfiguration),
        _ => throw new ArgumentException($"Invalid value \"{fileType}\".", nameof(fileType))
    };

    /// <summary>
    /// Constructs an importer that reads an Excel file from an existing data reader.
    /// </summary>
    /// <param name="reader">The existing data reader that wraps an Excel file.</param>
    public ExcelImporter(IExcelDataReader reader)
    {
        ThrowHelpers.ThrowIfNull(reader, nameof(reader));
        Reader = reader;
    }

    /// <summary>
    /// Cleans up resources associated with this class. This primarily involves disposing of
    /// the inner reader and any owned stream.
    /// </summary>
    public void Dispose()
    {
        Reader.Dispose();
        _ownedStream?.Dispose();
        _ownedStream = null;
    }

    /// <summary>
    /// Reads every row of the first sheet of an Excel or csv file and maps each row to
    /// <typeparamref name="T"/>. This is a shortcut for creating an importer, reading the first
    /// sheet and calling <see cref="ExcelSheet.ReadRows{T}()"/>.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <param name="path">A path to the file to read.</param>
    /// <param name="fileType">The type of file.</param>
    /// <param name="configure">Configures the importer before reading, for example to register a class map.</param>
    /// <returns>The mapped rows.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public static List<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(string path, ExcelImporterFileType fileType = ExcelImporterFileType.Excel, Action<ExcelImporterConfiguration>? configure = null)
    {
        using var importer = new ExcelImporter(path, fileType);
        return ReadRows<T>(importer, configure);
    }

    /// <summary>
    /// Reads every row of the first sheet of an Excel or csv stream and maps each row to
    /// <typeparamref name="T"/>. This is a shortcut for creating an importer, reading the first
    /// sheet and calling <see cref="ExcelSheet.ReadRows{T}()"/>. As with disposing an importer
    /// created from a stream, the stream is closed when reading finishes.
    /// </summary>
    /// <typeparam name="T">The type of the object to map each row to.</typeparam>
    /// <param name="stream">A stream containing the file to read.</param>
    /// <param name="fileType">The type of file.</param>
    /// <param name="configure">Configures the importer before reading, for example to register a class map.</param>
    /// <returns>The mapped rows.</returns>
    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    public static List<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Stream stream, ExcelImporterFileType fileType = ExcelImporterFileType.Excel, Action<ExcelImporterConfiguration>? configure = null)
    {
        using var importer = new ExcelImporter(stream, fileType);
        return ReadRows<T>(importer, configure);
    }

    [RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
    [RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
    private static List<T> ReadRows<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(ExcelImporter importer, Action<ExcelImporterConfiguration>? configure)
    {
        configure?.Invoke(importer.Configuration);
        return [.. importer.ReadSheet().ReadRows<T>()];
    }

    /// <summary>
    /// Reads each sheet in the document. Reading is reset at the end of enumeration.
    /// </summary>
    /// <returns>A lazily evaluated list of each sheet in the document.</returns>
    public IEnumerable<ExcelSheet> ReadSheets()
    {
        while (TryReadSheet(out var sheet))
        {
            yield return sheet;
        }

        ResetReader();
    }

    /// <summary>
    /// Reads the next sheet in the document. If no sheets have been read, then this reads the first sheet.
    /// The method throws if there are no more sheets in the document.
    /// </summary>
    /// <returns>The next sheet in the document.</returns>
    public ExcelSheet ReadSheet()
    {
        if (!TryReadSheet(out var sheet))
        {
            throw new ExcelMappingException("No more sheets.");
        }

        return sheet;
    }

    /// <summary>
    /// Finds and reads a sheet with the given name in the document.
    /// The method throws if the sheet does not exist.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read.</param>
    /// <returns>The sheet in the document with the given name.</returns>
    public ExcelSheet ReadSheet(string sheetName)
    {
        ThrowHelpers.ThrowIfNull(sheetName, nameof(sheetName));

        if (!TryReadSheet(sheetName, out var sheet))
        {
            throw new ExcelMappingException($"The sheet \"{sheetName}\" does not exist.");
        }

        return sheet;
    }

    /// <summary>
    /// Finds and reads a sheet at the given zero-based index in the document.
    /// The method throws if the index is invalid.
    /// </summary>
    /// <param name="sheetIndex">The index of the sheet to read.</param>
    /// <returns>The sheet in the document at the given zero-based index.</returns>
    public ExcelSheet ReadSheet(int sheetIndex)
    {
        if (!TryReadSheet(sheetIndex, out var sheet))
        {
            throw new ArgumentOutOfRangeException(nameof(sheetIndex), sheetIndex, $"The sheet index {sheetIndex} must be between 0 and {NumberOfSheets - 1}.");
        }

        return sheet;
    }

    /// <summary>
    /// Reads the next sheet in the document. If no sheets have been read, then this reads the first sheet.
    /// </summary>
    /// <param name="sheet">The next sheet in the document.</param>
    /// <returns>False if there are no more sheets in the document, else true.</returns>
    public bool TryReadSheet([NotNullWhen(true)] out ExcelSheet? sheet)
    {
        if (Reader.IsClosed)
        {
            throw new ExcelMappingException($"The underlying reader is closed.");
        }

        sheet = null;
        if (!RowSource.TryMoveToSheet(SheetIndex + 1))
        {
            return false;
        }

        SheetIndex++;
        sheet = new ExcelSheet(new Worksheet(Reader, SheetIndex, this));
        return true;
    }

    /// <summary>
    /// Finds and reads a sheet with the given name in the document.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read.</param>
    /// <param name="sheet">The sheet in the document with the given name.</param>
    /// <returns>True if the sheet was found, else false.</returns>
    public bool TryReadSheet(string sheetName, [NotNullWhen(true)] out ExcelSheet? sheet)
    {
        if (Reader.IsClosed)
        {
            throw new ExcelMappingException($"The underlying reader is closed.");
        }

        sheet = null;
        ResetReader();

        while (TryReadSheet(out ExcelSheet? currentSheet))
        {
            // Excel sheet names are case-insensitive.
            if (string.Equals(currentSheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                ResetReader();
                sheet = currentSheet;
                return true;
            }
        }

        ResetReader();
        return false;
    }

    /// <summary>
    /// Finds and reads a sheet at the given zero-based index in the document.
    /// </summary>
    /// <param name="sheetIndex">The zero-based index of the sheet to read.</param>
    /// <param name="sheet">The sheet in the document at the given zero-based index.</param>
    /// <returns>True if the sheet was found, else false.</returns>
    public bool TryReadSheet(int sheetIndex, [NotNullWhen(true)] out ExcelSheet? sheet)
    {
        if (Reader.IsClosed)
        {
            throw new ExcelMappingException($"The underlying reader is closed.");
        }

        ThrowHelpers.ThrowIfNegative(sheetIndex, nameof(sheetIndex));

        sheet = null;
        if (sheetIndex >= NumberOfSheets)
        {
            return false;
        }

        RowSource.TryMoveToSheet(sheetIndex);
        sheet = new ExcelSheet(new Worksheet(Reader, sheetIndex, this));
        ResetReader();
        return true;
    }

    // Start reading sheets in order from the first sheet again. The reader itself is moved when
    // it's next needed.
    private void ResetReader() => SheetIndex = -1;
}
