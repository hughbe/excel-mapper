using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;

namespace ExcelMapper.Tests;

public class ImporterOptionsTests
{
    // ---------- Reader configuration ----------

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

    // ---------- Trimming keeps typed values ----------

    [Fact]
    public void ReadRow_TrimmedDateCellWithCustomFormat_KeepsDate()
    {
        // The cell is a real date. Trimming used to turn it into culture-formatted text, which then
        // failed to parse with the custom format.
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTime>(c => c.Map(p => p).WithFormats("yyyy-MM-dd").WithTrim());

        var sheet = importer.ReadSheet();
        Assert.Equal(new DateTime(2017, 07, 19), sheet.ReadRow<DateTime>());
    }

    public class TrimmedValues
    {
        [ExcelTrimString]
        public string Name { get; set; } = "";

        [ExcelTrimString]
        public int Age { get; set; }
    }

    [Fact]
    public void ReadRows_TrimmedText_StillTrims()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\n\"  Alice  \",\" 30 \"\n");
        var row = importer.ReadSheet().ReadRows<TrimmedValues>().Single();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(30, row.Age);
    }

    // ---------- Ignoring separators in column names ----------

    public record Contact(string FirstName, string LastName, DateTime? DateOfBirth, string EMail);

    private const string ContactsCsv = "First Name,last_name,Date-Of-Birth,E.Mail\nAlice,Smith,,alice@example.org\n";

    [Fact]
    public void ReadRows_IgnoreSeparatorsInColumnNames_MatchesColumns()
    {
        using var importer = Helpers.GetCsvImporter(ContactsCsv);
        importer.Configuration.IgnoreSeparatorsInColumnNames = true;

        var row = importer.ReadSheet().ReadRows<Contact>().Single();
        Assert.Equal(new Contact("Alice", "Smith", null, "alice@example.org"), row);
    }

    [Fact]
    public void ReadRows_SeparatorsNotIgnored_ThrowsWithHint()
    {
        using var importer = Helpers.GetCsvImporter(ContactsCsv);
        var sheet = importer.ReadSheet();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Contact>().ToArray());
        Assert.Contains("for column \"FirstName\"", ex.Message);
        Assert.EndsWith("Did you mean column \"First Name\"? Use ExcelColumnName or WithColumnName, or set IgnoreSeparatorsInColumnNames in the importer's configuration.", ex.Message);
    }

    public class NameClass
    {
        public string FirstName { get; set; } = "";
    }

    [Fact]
    public void ReadRows_IgnoreSeparatorsWithExactMatch_PrefersExactMatch()
    {
        using var importer = Helpers.GetCsvImporter("First Name,FirstName\nloose,exact\n");
        importer.Configuration.IgnoreSeparatorsInColumnNames = true;
        Assert.Equal("exact", importer.ReadSheet().ReadRows<NameClass>().Single().FirstName);
    }

    [Fact]
    public void ReadRows_IgnoreSeparatorsWithFluentColumnName_Success()
    {
        using var importer = Helpers.GetCsvImporter("Id,first-name\n1,Alice\n");
        importer.Configuration.IgnoreSeparatorsInColumnNames = true;
        importer.Configuration.RegisterClassMap<NameClass>(c => c.Map(p => p.FirstName).WithColumnName("First Name"));
        Assert.Equal("Alice", importer.ReadSheet().ReadRows<NameClass>().Single().FirstName);
    }

    [Fact]
    public void ReadRows_NoSimilarColumn_ThrowsWithoutHint()
    {
        using var importer = Helpers.GetCsvImporter("Surname\nSmith\n");
        var sheet = importer.ReadSheet();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<NameClass>().ToArray());
        Assert.DoesNotContain("Did you mean", ex.Message);
    }

    [Fact]
    public void IgnoreSeparatorsInColumnNames_Default_IsFalse()
    {
        using var importer = Helpers.GetCsvImporter("a\n1\n");
        Assert.False(importer.Configuration.IgnoreSeparatorsInColumnNames);
    }
}
