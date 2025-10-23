using System.Diagnostics.CodeAnalysis;
using System.IO;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// An importer that reads the sheets in an Excel file or stream.
/// </summary>
public class ExcelImporter : IDisposable
{
    /// <summary>
    /// Gets the inner reader that the importer wraps. This is defined by the ExcelDataReader library.
    /// </summary>
    public IExcelDataReader Reader { get; }

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

    private int SheetIndex { get; set; } = -1;

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
    public ExcelImporter(string path, ExcelImporterFileType fileType)
    {
        ArgumentNullException.ThrowIfNull(path);

        Stream? stream = null;
        try
        {
            stream = File.OpenRead(path);
            Reader = fileType switch
            {
                ExcelImporterFileType.Excel => ExcelReaderFactory.CreateReader(stream),
                ExcelImporterFileType.Csv => ExcelReaderFactory.CreateCsvReader(stream),
                _ => throw new ArgumentException($"Invalid value \"{fileType}\".", nameof(fileType))
            };
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
    public ExcelImporter(Stream stream, ExcelImporterFileType fileType)
    {
        ArgumentNullException.ThrowIfNull(stream);

        Reader = fileType switch
        {
            ExcelImporterFileType.Excel => ExcelReaderFactory.CreateReader(stream),
            ExcelImporterFileType.Csv => ExcelReaderFactory.CreateCsvReader(stream),
            _ => throw new ArgumentException($"Invalid value \"{fileType}\".", nameof(fileType))
        };
    }

    /// <summary>
    /// Constructs an importer that reads an Excel file from an existing data reader.
    /// </summary>
    /// <param name="reader">The existing data reader that wraps an Excel file.</param>
    public ExcelImporter(IExcelDataReader reader)
    {
        ArgumentNullException.ThrowIfNull(reader);
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
        ArgumentNullException.ThrowIfNull(sheetName);

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

        if (SheetIndex != -1)
        {
            if (!Reader.NextResult())
            {
                return false;
            }
        }

        SheetIndex++;
        sheet = new ExcelSheet(Reader, SheetIndex, this);
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
            if (currentSheet.Name == sheetName)
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

        ArgumentOutOfRangeException.ThrowIfNegative(sheetIndex);

        sheet = null;
        if (sheetIndex >= NumberOfSheets)
        {
            return false;
        }

        ResetReader();
        for (int i = 0; i < sheetIndex; i++)
        {
            Reader.NextResult();
        }

        sheet = new ExcelSheet(Reader, sheetIndex, this);
        ResetReader();
        return true;
    }

    private void ResetReader()
    {
        Reader.Reset();
        SheetIndex = -1;
    }

    internal void MoveToSheet(ExcelSheet sheet)
    {
        // Already on the sheet.
        if (SheetIndex == sheet.Index)
        {
            return;
        }

        // Read up to the current sheet.
        Reader.Reset();
        for (int i = 0; i < sheet.Index; i++)
        {
            Reader.NextResult();
        }

        // If the header has already been read, skip past it.
        if (sheet.HasHeading && sheet.Heading != null)
        {
            sheet.ReadPastHeading();
        }

        // Read up to the current row.
        for (int i = 0; i <= sheet.CurrentRowIndex; i++)
        {
            Reader.Read();
        }

        SheetIndex = sheet.Index;
    }
}
