using System.Text;
using System.IO;
using System.Linq;
using ExcelDataReader;
using ExcelDataReader.Exceptions;

namespace ExcelMapper.Tests;

public class ExcelImporterTests
{
    [Fact]
    public void Ctor_String_Xlsx()
    {
        var filePath = Helpers.GetResourcePath("Primitives.xlsx");
        using var importer = new ExcelImporter(filePath);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }

    [Fact]
    public void Ctor_String_FileOpenForWriting_Success()
    {
        // As Excel does with a workbook it has open.
        var filePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
        File.Copy(Helpers.GetResourcePath("Primitives.xlsx"), filePath);
        try
        {
            using (new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            {
                // Only Windows enforces sharing between readers and writers, so only there does a
                // file opened just for reading fail.
                if (OperatingSystem.IsWindows())
                {
                    Assert.Throws<IOException>(() => File.OpenRead(filePath).Dispose());
                }

                using var importer = new ExcelImporter(filePath);
                Assert.Equal("Primitives", importer.ReadSheet().Name);
            }
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Ctor_String_Xlsm()
    {
        var filePath = Helpers.GetResourcePath("Primitives.xlsm");
        using var importer = new ExcelImporter(filePath);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }

    [Fact]
    public void Ctor_String_Xls()
    {
        var filePath = Helpers.GetResourcePath("Primitives.xls");
        using var importer = new ExcelImporter(filePath);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }

    [Fact]
    public void Ctor_String_Xlsb()
    {
        var filePath = Helpers.GetResourcePath("Primitives.xlsb");
        using var importer = new ExcelImporter(filePath);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }

    [Fact]
    public void Ctor_String_OpenXml()
    {
        var filePath = Helpers.GetResourcePath("OpenXML.xlsx");
        using var importer = new ExcelImporter(filePath);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }

    [Fact]
    public void Ctor_String_ExcelImporterFileType_Excel()
    {
        var filePath = Helpers.GetResourcePath("Primitives.xlsx");
        using var importer = new ExcelImporter(filePath, ExcelImporterFileType.Excel);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }
    
    [Fact]
    public void Ctor_String_ExcelImporterFileType_Csv()
    {
        var filePath = Helpers.GetResourcePath("Primitives.csv");
        using var importer = new ExcelImporter(filePath, ExcelImporterFileType.Csv);
        Assert.Equal(string.Empty, importer.ReadSheet().Name);
        Assert.Equal(1, importer.NumberOfSheets);
    }

    [Fact]
    public void Ctor_NullPath_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("path", () => new ExcelImporter((string)null!));
        Assert.Throws<ArgumentNullException>("path", () => new ExcelImporter((string)null!, ExcelImporterFileType.Csv));
    }

    [Fact]
    public void Ctor_EmptyPath_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("path", () => new ExcelImporter(string.Empty));
        Assert.Throws<ArgumentException>("path", () => new ExcelImporter(string.Empty, ExcelImporterFileType.Csv));
    }

    [Fact]
    public void Ctor_NoSuchFile_ThrowsFileNotFoundException()
    {
        var filePath = Helpers.GetResourcePath("NoSuchFile.xlsx");
        Assert.Throws<FileNotFoundException>(() => new ExcelImporter(filePath));
        Assert.Throws<FileNotFoundException>(() => new ExcelImporter(filePath, ExcelImporterFileType.Csv));
    }

    [Fact]
    public void Ctor_CsvFile_ReadsCsv()
    {
        var filePath = Helpers.GetResourcePath("Primitives.csv");
        using var importer = new ExcelImporter(filePath);
        Assert.Equal(1, importer.NumberOfSheets);
        Assert.NotEmpty(importer.ReadSheet().ReadRows<string[]>());
    }

    [Fact]
    public void Ctor_CsvFileAsExcel_ThrowsHeaderException()
    {
        var filePath = Helpers.GetResourcePath("Primitives.csv");
        Assert.Throws<HeaderException>(() => new ExcelImporter(filePath, ExcelImporterFileType.Excel));
    }

    [Theory]
    [InlineData(".csv")]
    [InlineData(".CSV")]
    [InlineData(".tsv")]
    [InlineData(".txt")]
    public void Ctor_CsvExtension_ReadsCsv(string extension)
    {
        var filePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + extension);
        File.WriteAllText(filePath, "Name\tAge\nAlice\t34\n");
        try
        {
            using var importer = new ExcelImporter(filePath);
            var row = Assert.Single(importer.ReadSheet().ReadRows<Dictionary<string, string>>());
            Assert.Equal("34", row["Age"]);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Ctor_OdsFile_ThrowsHeaderException()
    {
        var filePath = Helpers.GetResourcePath("Primitives.ods");
        Assert.Throws<HeaderException>(() => new ExcelImporter(filePath));
        Assert.Throws<HeaderException>(() => new ExcelImporter(filePath, ExcelImporterFileType.Excel));
    }

    [Fact]
    public void Ctor_Stream_Xlsx()
    {
        using var stream = Helpers.GetResource("Primitives.xlsx");
        using var importer = new ExcelImporter(stream);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }
    
    [Fact]
    public void Ctor_Stream_Xlsm()
    {
        using var stream = Helpers.GetResource("Primitives.xlsm");
        using var importer = new ExcelImporter(stream);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }
    
    [Fact]
    public void Ctor_Stream_Xls()
    {
        using var stream = Helpers.GetResource("Primitives.xls");
        using var importer = new ExcelImporter(stream);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }
    
    [Fact]
    public void Ctor_Stream_Xlsb()
    {
        using var stream = Helpers.GetResource("Primitives.xlsb");
        using var importer = new ExcelImporter(stream);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }
    
    [Fact]
    public void Ctor_Stream_OpenXml()
    {
        using var stream = Helpers.GetResource("OpenXML.xlsx");
        using var importer = new ExcelImporter(stream);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }

    [Fact]
    public void Ctor_Stream_ExcelImporterFileType_Excel()
    {
        using var stream = Helpers.GetResource("Primitives.xlsx");
        using var importer = new ExcelImporter(stream, ExcelImporterFileType.Excel);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }
    
    [Fact]
    public void Ctor_Stream_ExcelImporterFileType_Csv()
    {
        using var stream = Helpers.GetResource("Primitives.csv");
        using var importer = new ExcelImporter(stream, ExcelImporterFileType.Csv);
        Assert.Equal(string.Empty, importer.ReadSheet().Name);
        Assert.Equal(1, importer.NumberOfSheets);
    }

    [Fact]
    public void Ctor_EmptyStream_ThrowsHeaderException()
    {
        using var stream = new MemoryStream();
        Assert.Throws<HeaderException>(() => new ExcelImporter(stream));
    }

    [Fact]
    public void Ctor_NullStream_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("stream", () => new ExcelImporter((Stream)null!));
        Assert.Throws<ArgumentNullException>("stream", () => new ExcelImporter((Stream)null!, ExcelImporterFileType.Csv));
    }

    [Theory]
    [InlineData(ExcelImporterFileType.Excel - 1)]
    [InlineData(ExcelImporterFileType.Csv + 1)]
    public void Ctor_InvalidFileType_ThrowsArgumentException(ExcelImporterFileType fileType)
    {
        using var stream = Helpers.GetResource("Strings.xlsx");
        Assert.Throws<ArgumentException>("fileType", () => new ExcelImporter(Helpers.GetResourcePath("Strings.xlsx"), fileType));
        Assert.Throws<ArgumentException>("fileType", () => new ExcelImporter(stream, fileType));
    }

    [Fact]
    public void Ctor_CsvStream_ThrowsHeaderException()
    {
        using var stream = Helpers.GetResource("Primitives.csv");
        Assert.Throws<HeaderException>(() => new ExcelImporter(stream));
        Assert.Throws<HeaderException>(() => new ExcelImporter(stream, ExcelImporterFileType.Excel));
    }

    [Fact]
    public void Ctor_OdsStream_ThrowsHeaderException()
    {
        using var stream = Helpers.GetResource("Primitives.ods");
        Assert.Throws<HeaderException>(() => new ExcelImporter(stream));
        Assert.Throws<HeaderException>(() => new ExcelImporter(stream, ExcelImporterFileType.Excel));
    }

    [Fact]
    public void Ctor_IExcelDataReader()
    {
        using var stream = Helpers.GetResource("Primitives.xlsx");
        using var reader = ExcelReaderFactory.CreateReader(stream);
        using var importer = new ExcelImporter(reader);
        Assert.Equal("Primitives", importer.ReadSheet().Name);
        Assert.Equal(3, importer.NumberOfSheets);
    }

    [Fact]
    public void Ctor_NullReader_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("reader", () => new ExcelImporter((IExcelDataReader)null!));
    }

    [Fact]
    public void ReadSheets_Invoke_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        ExcelSheet[] sheets = [.. importer.ReadSheets()];
        Assert.Equal(["Primitives", "Empty", "Third Sheet"], sheets.Select(sheet => sheet.Name));
        Assert.Equal([true, true, true], sheets.Select(sheet => sheet.HasHeading));
        Assert.Equal([null, null, null], sheets.Select(sheet => sheet.Heading));
    }

    [Fact]
    public void ReadSheets_InvokeReadSheetsAfter_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        ExcelSheet[] sheets = [.. importer.ReadSheets()];
        Assert.Equal(["Primitives", "Empty", "Third Sheet"], sheets.Select(sheet => sheet.Name));
        Assert.Equal([0, 1, 2], sheets.Select(sheet => sheet.Index));
        Assert.Equal([true, true, true], sheets.Select(sheet => sheet.HasHeading));
        Assert.Equal([null, null, null], sheets.Select(sheet => sheet.Heading));

        sheets = [.. importer.ReadSheets()];
        Assert.Equal(["Primitives", "Empty", "Third Sheet"], sheets.Select(sheet => sheet.Name));
        Assert.Equal([0, 1, 2], sheets.Select(sheet => sheet.Index));
        Assert.Equal([true, true, true], sheets.Select(sheet => sheet.HasHeading));
        Assert.Equal([null, null, null], sheets.Select(sheet => sheet.Heading));
    }

    [Fact]
    public void ReadSheets_InvokeReadSheetAfter_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        ExcelSheet[] sheets = [.. importer.ReadSheets()];
        Assert.Equal(["Primitives", "Empty", "Third Sheet"], sheets.Select(sheet => sheet.Name));
        Assert.Equal([0, 1, 2], sheets.Select(sheet => sheet.Index));
        Assert.Equal([true, true, true], sheets.Select(sheet => sheet.HasHeading));
        Assert.Equal([null, null, null], sheets.Select(sheet => sheet.Heading));

        ExcelSheet nextSheet = importer.ReadSheet();
        Assert.Equal("Primitives", nextSheet.Name);
        Assert.Equal(0, nextSheet.Index);
        Assert.True(nextSheet.HasHeading);
        Assert.Null(nextSheet.Heading);
    }

    [Fact]
    public void ReadSheets_LeftEarly_ReadSheetStartsFromFirstSheet()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        foreach (var sheet in importer.ReadSheets())
        {
            Assert.Equal("Primitives", sheet.Name);
            break;
        }

        Assert.Equal("Primitives", importer.ReadSheet().Name);
    }

    [Fact]
    public void ReadSheet_AllSheets_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Equal("Primitives", sheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(0, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        sheet = importer.ReadSheet();
        Assert.Equal("Empty", sheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(1, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        sheet = importer.ReadSheet();
        Assert.Equal("Third Sheet", sheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(2, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadSheet_NoMoreSheets_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.ReadSheet();
        importer.ReadSheet();
        importer.ReadSheet();

        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet());
    }

    [Fact]
    public void ReadSheet_ReaderClosedByCaller_ThrowsObjectDisposedException()
    {
        using var stream = Helpers.GetResource("Strings.xlsx");
        using var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        using var importer = new ExcelImporter(reader);
        reader.Close();

        var ex = Assert.Throws<ObjectDisposedException>(() => importer.ReadSheet());
        Assert.Equal(nameof(ExcelImporter), ex.ObjectName);
        Assert.StartsWith("The importer has been disposed, or the reader it was created with has been closed.", ex.Message);
    }

    [Fact]
    public void ReadSheet_Closed_ThrowsObjectDisposedException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Dispose();

        var ex = Assert.Throws<ObjectDisposedException>(() => importer.ReadSheet());
        Assert.Equal(nameof(ExcelImporter), ex.ObjectName);
    }

    [Fact]
    public void TryReadSheet_AllSheets_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.NotNull(sheet1);
        Assert.Equal("Primitives", sheet1!.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet1.Visibility);
        Assert.Equal(0, sheet1.Index);
        Assert.True(sheet1.HasHeading);
        Assert.Null(sheet1.Heading);
        Assert.Equal(-1, sheet1.CurrentRowIndex);

        Assert.True(importer.TryReadSheet(out var sheet2));
        Assert.NotNull(sheet2);
        Assert.Equal("Empty", sheet2!.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet2.Visibility);
        Assert.Equal(1, sheet2.Index);
        Assert.True(sheet2.HasHeading);
        Assert.Null(sheet2.Heading);
        Assert.Equal(-1, sheet2.CurrentRowIndex);

        Assert.True(importer.TryReadSheet(out var sheet3));
        Assert.NotNull(sheet3);
        Assert.Equal("Third Sheet", sheet3!.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet3.Visibility);
        Assert.Equal(2, sheet3.Index);
        Assert.True(sheet3.HasHeading);
        Assert.Null(sheet3.Heading);
        Assert.Equal(-1, sheet3.CurrentRowIndex);
    }

    [Fact]
    public void TryReadSheet_NoMoreSheets_ReturnsFalse()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.ReadSheet();
        importer.ReadSheet();
        importer.ReadSheet();

        Assert.False(importer.TryReadSheet(out var sheet));
        Assert.Null(sheet);
    }

    [Fact]
    public void TryReadSheet_Closed_ThrowsObjectDisposedException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Dispose();

        var ex = Assert.Throws<ObjectDisposedException>(() => importer.TryReadSheet(out var sheet));
        Assert.Equal(nameof(ExcelImporter), ex.ObjectName);
    }

    [Fact]
    public void ReadSheet_SheetNameExistsNotAlreadyRead_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet("Empty");
        Assert.Equal("Empty", sheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(1, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        // Reading a named sheet should reset the reader after finding the column.
        ExcelSheet nextSheet = importer.ReadSheet();
        Assert.Equal("Primitives", nextSheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
        Assert.Equal(0, nextSheet.Index);
        Assert.True(nextSheet.HasHeading);
        Assert.Null(nextSheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadSheet_SheetNameExistsAlreadyRead_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.ReadSheet();
        importer.ReadSheet();
        importer.ReadSheet();

        // Reading a named sheet should reset the reader before finding the sheet.
        var sheet = importer.ReadSheet("Empty");
        Assert.Equal("Empty", sheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(1, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        // Reading a named sheet should reset the reader after finding the column.
        ExcelSheet nextSheet = importer.ReadSheet();
        Assert.Equal("Primitives", nextSheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
        Assert.Equal(0, nextSheet.Index);
        Assert.True(nextSheet.HasHeading);
        Assert.Null(nextSheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadSheet_NullSheetName_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Throws<ArgumentNullException>("sheetName", () => importer.ReadSheet(null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData(" Empty ")]
    [InlineData("invalid")]
    [InlineData("  \r \t  ")]
    public void ReadSheet_NoSuchSheet_ThrowsExcelMappingException(string sheetName)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet(sheetName));
    }

    [Fact]
    public void ReadSheet_StringClosed_ThrowsObjectDisposedException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Dispose();

        var ex = Assert.Throws<ObjectDisposedException>(() => importer.ReadSheet("Sheet"));
        Assert.Equal(nameof(ExcelImporter), ex.ObjectName);
    }

    [Fact]
    public void TryReadSheet_SheetNameExistsNotAlreadyRead_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet("Empty", out var sheet));
        Assert.NotNull(sheet);
        Assert.Equal("Empty", sheet!.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(1, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        // Reading a named sheet should reset the reader after finding the column.
        ExcelSheet nextSheet = importer.ReadSheet();
        Assert.Equal("Primitives", nextSheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
        Assert.Equal(0, nextSheet.Index);
        Assert.True(nextSheet.HasHeading);
        Assert.Null(nextSheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);
    }

    [Fact]
    public void TryReadSheet_SheetNameExistsAlreadyRead_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.ReadSheet();
        importer.ReadSheet();
        importer.ReadSheet();

        // Reading a named sheet should reset the reader before finding the sheet.
        Assert.True(importer.TryReadSheet("Empty", out var sheet));
        Assert.NotNull(sheet);
        Assert.Equal("Empty", sheet!.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(1, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        // Reading a named sheet should reset the reader after finding the column.
        ExcelSheet nextSheet = importer.ReadSheet();
        Assert.Equal("Primitives", nextSheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
        Assert.Equal(0, nextSheet.Index);
        Assert.True(nextSheet.HasHeading);
        Assert.Null(nextSheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);
    }

    [Theory]
    [InlineData("empty")]
    [InlineData("EMPTY")]
    [InlineData("eMpTy")]
    public void ReadSheet_SheetNameDifferentCase_ReturnsSheet(string sheetName)
    {
        // Excel sheet names are case-insensitive.
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet(sheetName);
        Assert.Equal("Empty", sheet.Name);
        Assert.Equal(1, sheet.Index);

        Assert.True(importer.TryReadSheet(sheetName, out var trySheet));
        Assert.Equal("Empty", trySheet.Name);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" Empty ")]
    [InlineData("invalid")]
    [InlineData("  \r \t  ")]
    public void TryReadSheet_NoSuchSheet_ReturnsFalse(string? sheetName)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.False(importer.TryReadSheet(sheetName!, out var sheet));
        Assert.Null(sheet);
    }

    [Fact]
    public void TryReadSheet_StringClosed_ThrowsObjectDisposedException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Dispose();

        var ex = Assert.Throws<ObjectDisposedException>(() => importer.TryReadSheet("Sheet", out var sheet));
        Assert.Equal(nameof(ExcelImporter), ex.ObjectName);
    }

    [Fact]
    public void ReadSheet_SheetIndexExistsNotAlreadyRead_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet(1);
        Assert.Equal("Empty", sheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(1, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        // Reading a named sheet should reset the reader after finding the column.
        ExcelSheet nextSheet = importer.ReadSheet();
        Assert.Equal("Primitives", nextSheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
        Assert.Equal(0, nextSheet.Index);
        Assert.True(nextSheet.HasHeading);
        Assert.Null(nextSheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadSheet_SheetIndexExistsAlreadyRead_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.ReadSheet();
        importer.ReadSheet();
        importer.ReadSheet();

        // Reading a named sheet should reset the reader before finding the sheet.
        var sheet = importer.ReadSheet(1);
        Assert.Equal("Empty", sheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(1, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        // Reading a named sheet should reset the reader after finding the column.
        ExcelSheet nextSheet = importer.ReadSheet();
        Assert.Equal("Primitives", nextSheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
        Assert.Equal(0, nextSheet.Index);
        Assert.True(nextSheet.HasHeading);
        Assert.Null(nextSheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(3)]
    public void ReadSheet_InvalidSheetIndex_ThrowsArgumentOutOfRangeException(int sheetIndex)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Throws<ArgumentOutOfRangeException>("sheetIndex", () => importer.ReadSheet(sheetIndex));
    }

    [Fact]
    public void ReadSheet_IntClosed_ThrowsObjectDisposedException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Dispose();

        var ex = Assert.Throws<ObjectDisposedException>(() => importer.ReadSheet(0));
        Assert.Equal(nameof(ExcelImporter), ex.ObjectName);
    }

    [Fact]
    public void TryReadSheet_SheetIndexExistsNotAlreadyRead_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(1, out var sheet));
        Assert.NotNull(sheet);
        Assert.Equal("Empty", sheet!.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(1, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        // Reading a named sheet should reset the reader after finding the column.
        ExcelSheet nextSheet = importer.ReadSheet();
        Assert.Equal("Primitives", nextSheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
        Assert.Equal(0, nextSheet.Index);
        Assert.True(nextSheet.HasHeading);
        Assert.Null(nextSheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);
    }

    [Fact]
    public void TryReadSheet_SheetIndexExistsAlreadyRead_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.ReadSheet();
        importer.ReadSheet();
        importer.ReadSheet();

        // Reading a named sheet should reset the reader before finding the sheet.
        Assert.True(importer.TryReadSheet(1, out var sheet));
        Assert.NotNull(sheet);
        Assert.Equal("Empty", sheet!.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
        Assert.Equal(1, sheet.Index);
        Assert.True(sheet.HasHeading);
        Assert.Null(sheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);

        // Reading a named sheet should reset the reader after finding the column.
        ExcelSheet nextSheet = importer.ReadSheet();
        Assert.Equal("Primitives", nextSheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, nextSheet.Visibility);
        Assert.Equal(0, nextSheet.Index);
        Assert.True(nextSheet.HasHeading);
        Assert.Null(nextSheet.Heading);
        Assert.Equal(-1, sheet.CurrentRowIndex);
    }

    [Theory]
    [InlineData(3)]
    [InlineData(4)]
    public void TryReadSheet_InvalidSheetIndex_ReturnsFalse(int sheetIndex)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.False(importer.TryReadSheet(sheetIndex, out var sheet));
        Assert.Null(sheet);
    }

    [Fact]
    public void TryReadSheet_NegativeSheetIndex_ThrowsArgumentOutOfRangeException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Throws<ArgumentOutOfRangeException>("sheetIndex", () => importer.TryReadSheet(-1, out var sheet));
    }

    [Fact]
    public void TryReadSheet_IntClosed_ThrowsObjectDisposedException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Dispose();

        var ex = Assert.Throws<ObjectDisposedException>(() => importer.TryReadSheet(0, out var sheet));
        Assert.Equal(nameof(ExcelImporter), ex.ObjectName);
    }

    [Fact]
    public void MultipleSheets_ReadFirstHeadingThenFirstRowsThenLastHeadingThenLastRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.ReadHeading();

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Null(sheet1Row3.StringValue);

        sheet3!.ReadHeading();

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s3", sheet3Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_ReadFirstHeadingThenLastHeadingThenFirstRowsThenLastRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.ReadHeading();
        sheet3!.ReadHeading();

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Null(sheet1Row3.StringValue);

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s3", sheet3Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_ReadFirstHeadingThenLastHeadingThenLastRowsThenFirstRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.ReadHeading();
        sheet3!.ReadHeading();

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s3", sheet3Row3.StringValue);

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Null(sheet1Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_ReadLastHeadingThenLastRowsThenFirstHeadingThenFirstRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet3!.ReadHeading();

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s3", sheet3Row3.StringValue);

        sheet1!.ReadHeading();

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Null(sheet1Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_ReadLastHeadingThenFirstHeadingThenFirstRowsThenLastRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet3!.ReadHeading();
        sheet1!.ReadHeading();

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Null(sheet1Row3.StringValue);

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s3", sheet3Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_ReadLastHeadingThenFirstHeadingThenLastRowsThenFirstRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet3!.ReadHeading();
        sheet1!.ReadHeading();

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s3", sheet3Row3.StringValue);

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Null(sheet1Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_ReadFirstRowThenLastRowThenFirstRowThenLastRow_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.ReadHeading();
        sheet3!.ReadHeading();

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row2.StringValue);

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Null(sheet1Row3.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s3", sheet3Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_ReadFirstRowThenLastRowThenLastRowThenFirstRow_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.ReadHeading();
        sheet3!.ReadHeading();

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row1.StringValue);

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row2.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Null(sheet1Row3.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s3", sheet3Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_ReadLastRowThenFirstRowThenLastRowThenFirstRow_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.ReadHeading();
        sheet3!.ReadHeading();

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row1.StringValue);

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row2.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s3", sheet3Row3.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Null(sheet1Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_ReadLastRowThenFirstRowThenLastRowThenLastRowThenFirstRow_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.ReadHeading();
        sheet3!.ReadHeading();

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row1.StringValue);

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row2.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s3", sheet3Row3.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Null(sheet1Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_NoHeadingReadFirstRowsThenLastRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.HasHeading = false;
        sheet3!.HasHeading = false;

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row3.StringValue);

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet3Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_NoHeadingReadLastRowsThenFirstRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.HasHeading = false;
        sheet3!.HasHeading = false;

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet3Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row3.StringValue);

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_NoHeadingReadFirstRowsThenLastRowsThenFirstRowsThenLastRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.HasHeading = false;
        sheet3!.HasHeading = false;

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet1Row1.StringValue);

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet3Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row2.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row3.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_NoHeadingReadFirstRowsThenLastRowsThenLastRowsThenFirstRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.HasHeading = false;
        sheet3!.HasHeading = false;

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet1Row1.StringValue);

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet3Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row2.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row2.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row3.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_NoHeadingReadLastRowsThenFirstRowsThenFirstRowsThenLastRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.HasHeading = false;
        sheet3!.HasHeading = false;

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet3Row1.StringValue);

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet1Row1.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row2.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row3.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row3.StringValue);
    }

    [Fact]
    public void MultipleSheets_NoHeadingReadLastRowsThenFirstRowsThenLastRowsThenFirstRows_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestClassMapColumnIndex>();

        Assert.True(importer.TryReadSheet(out var sheet1));
        Assert.True(importer.TryReadSheet(out _));
        Assert.True(importer.TryReadSheet(out var sheet3));

        sheet1!.HasHeading = false;
        sheet3!.HasHeading = false;

        var sheet3Row1 = sheet3.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet3Row1.StringValue);

        var sheet1Row1 = sheet1.ReadRow<TestClass>();
        Assert.Equal("StringValue", sheet1Row1.StringValue);

        var sheet3Row2 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s1", sheet3Row2.StringValue);

        var sheet1Row2 = sheet1.ReadRow<TestClass>();
        Assert.Equal("a", sheet1Row2.StringValue);

        var sheet3Row3 = sheet3.ReadRow<TestClass>();
        Assert.Equal("s2", sheet3Row3.StringValue);

        var sheet1Row3 = sheet1.ReadRow<TestClass>();
        Assert.Equal("b", sheet1Row3.StringValue);
    }

    private class TestClassMapColumnIndex : ExcelClassMap<TestClass>
    {
        public TestClassMapColumnIndex()
        {
            Map(value => value.StringValue)
                .WithColumnIndex(1);
        }
    }

    private class TestClass
    {
        public string? StringValue { get; set; }
    }

    [Fact]
    public void Ctor_StreamCsvReaderConfiguration_UsesSeparator()
    {
        var configuration = new ExcelReaderConfiguration { AutodetectSeparators = [';'] };
        using var importer = new ExcelImporter(new MemoryStream(Encoding.UTF8.GetBytes("a,b;c\n1,2;3\n")), ExcelImporterFileType.Csv, configuration);

        var sheet = importer.ReadSheet();
        Assert.Equal(["a,b", "c"], sheet.ReadHeading().ColumnNames);
        Assert.Equal(["1,2", "3"], sheet.ReadRow<string[]>());
    }

    [Fact]
    public void Ctor_StreamCsvNullReaderConfiguration_UsesDefaults()
    {
        using var importer = new ExcelImporter(new MemoryStream(Encoding.UTF8.GetBytes("a,b\n1,2\n")), ExcelImporterFileType.Csv, null);
        Assert.Equal(["a", "b"], importer.ReadSheet().ReadHeading().ColumnNames);
    }

    [Fact]
    public void Ctor_PathReaderConfiguration_Success()
    {
        using var importer = new ExcelImporter(Helpers.GetResourcePath("Primitives.xlsx"), ExcelImporterFileType.Excel, new ExcelReaderConfiguration());
        Assert.Equal("Primitives", importer.ReadSheet().Name);
    }

    [Fact]
    public void Ctor_StreamExcelReaderConfiguration_Success()
    {
        using var stream = Helpers.GetResource("Primitives.xlsx");
        using var importer = new ExcelImporter(stream, ExcelImporterFileType.Excel, new ExcelReaderConfiguration());
        Assert.Equal(3, importer.NumberOfSheets);
    }

    [Fact]
    public void Ctor_InvalidFileTypeReaderConfiguration_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("fileType", () => new ExcelImporter(new MemoryStream(), (ExcelImporterFileType)99, null));
        Assert.Throws<ArgumentException>("fileType", () => new ExcelImporter(Helpers.GetResourcePath("Primitives.xlsx"), (ExcelImporterFileType)99, null));
    }

    [Fact]
    public void Ctor_Any_RegistersCodePageEncodings()
    {
        using var importer = Helpers.GetCsvImporter("a\n1\n");

        // Windows-1252 is used by many .xls and CSV files and isn't available on .NET by default.
        Assert.Equal(1252, Encoding.GetEncoding(1252).CodePage);
    }

    [Fact]
    public void StaticReadRows_Path_ReadsFirstSheet()
    {
        var rows = ExcelImporter.ReadRows<string[]>(Helpers.GetResourcePath("Primitives.xlsx"));
        Assert.NotEmpty(rows);
    }

    [Fact]
    public void StaticReadRows_PathCsv_Success()
    {
        var rows = ExcelImporter.ReadRows<string[]>(Helpers.GetResourcePath("Primitives.csv"), ExcelImporterFileType.Csv);
        Assert.NotEmpty(rows);
    }

    [Fact]
    public void StaticReadRows_StreamWithConfiguration_Success()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Full Name,Age\nAlice,30\nBob,25\n"));
        var rows = ExcelImporter.ReadRows<StaticReadRowsPerson>(stream, ExcelImporterFileType.Csv, c =>
            c.RegisterClassMap<StaticReadRowsPerson>(m =>
            {
                m.Map(p => p.Name).WithColumnName("Full Name");
                m.Map(p => p.Age);
            }));

        Assert.Equal([new StaticReadRowsPerson("Alice", 30), new StaticReadRowsPerson("Bob", 25)], rows);
    }

    [Fact]
    public void StaticReadRows_Stream_ClosesStream()
    {
        // Matches disposing an importer created from a stream: ExcelDataReader closes the stream.
        var stream = new MemoryStream(Encoding.UTF8.GetBytes("Name,Age\nAlice,30\n"));
        Assert.Single(ExcelImporter.ReadRows<StaticReadRowsPerson>(stream, ExcelImporterFileType.Csv));
        Assert.False(stream.CanRead);
    }

    [Fact]
    public void StaticReadRows_PathWithConfiguration_Success()
    {
        // An Excel file, configured without giving the file type.
        var rows = ExcelImporter.ReadRows<string[]>(Helpers.GetResourcePath("Primitives.xlsx"), c => c.SkipBlankLines = true);
        Assert.NotEmpty(rows);
    }

    [Fact]
    public void StaticReadRows_StreamWithConfigurationOnly_ReadsExcel()
    {
        using var stream = File.OpenRead(Helpers.GetResourcePath("Primitives.xlsx"));
        var configured = false;
        var rows = ExcelImporter.ReadRows<string[]>(stream, c => configured = true);
        Assert.NotEmpty(rows);
        Assert.True(configured);
    }

    [Fact]
    public void StaticReadRows_PathAndFileTypeWithConfiguration_Success()
    {
        var rows = ExcelImporter.ReadRows<string[]>(Helpers.GetResourcePath("Primitives.csv"), ExcelImporterFileType.Csv, c => c.SkipBlankLines = true);
        Assert.NotEmpty(rows);
    }

    [Fact]
    public void StaticReadRows_StreamDefault_ReadsExcel()
    {
        using var stream = File.OpenRead(Helpers.GetResourcePath("Primitives.xlsx"));
        Assert.NotEmpty(ExcelImporter.ReadRows<string[]>(stream));
    }

    [Fact]
    public void StaticReadRows_NullArguments_Throw()
    {
        Assert.Throws<ArgumentNullException>("path", () => ExcelImporter.ReadRows<StaticReadRowsPerson>((string)null!));
        Assert.Throws<ArgumentNullException>("stream", () => ExcelImporter.ReadRows<StaticReadRowsPerson>((Stream)null!));
        Assert.Throws<ArgumentNullException>("configure", () => ExcelImporter.ReadRows<StaticReadRowsPerson>(Helpers.GetResourcePath("Primitives.xlsx"), (Action<ExcelImporterConfiguration>)null!));
        Assert.Throws<ArgumentNullException>("configure", () => ExcelImporter.ReadRows<StaticReadRowsPerson>(Helpers.GetResourcePath("Primitives.csv"), ExcelImporterFileType.Csv, null!));
        Assert.Throws<ArgumentNullException>("configure", () => ExcelImporter.ReadRows<StaticReadRowsPerson>(new MemoryStream(), (Action<ExcelImporterConfiguration>)null!));
        Assert.Throws<ArgumentNullException>("configure", () => ExcelImporter.ReadRows<StaticReadRowsPerson>(new MemoryStream(), ExcelImporterFileType.Csv, null!));
    }

    public record StaticReadRowsPerson(string Name, int Age);
}
