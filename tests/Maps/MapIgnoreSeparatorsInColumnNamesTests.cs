using System.Linq;

namespace ExcelMapper.Tests;

public class MapIgnoreSeparatorsInColumnNamesTests
{
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
}
