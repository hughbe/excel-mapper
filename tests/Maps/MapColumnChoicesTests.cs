using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelMapper.Tests;

/// <summary>
/// Members that can be read from one of several columns, or from columns that may be missing.
/// </summary>
public class MapColumnChoicesTests
{
    [Theory]
    [InlineData("Full Name,Age\na,1\n", "a")]
    [InlineData("Name,Age\nb,1\n", "b")]
    [InlineData("Full Name,NAME\nc,d\n", "c")]
    public void ReadRows_SeveralColumnNameAttributes_ReadsFirstColumnFound(string csv, string expected)
    {
        using var importer = Helpers.GetCsvImporter(csv);
        Assert.Equal(expected, importer.ReadSheet().ReadRows<SeveralNames>().Single().Name);
    }

    [Fact]
    public void ReadRows_SeveralColumnNameAttributes_EachUsesItsComparison()
    {
        // "Full Name" is compared ordinally, so "FULL NAME" only matches the second, case-insensitive name.
        using var importer = Helpers.GetCsvImporter("FULL NAME\na\n");
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<OrdinalNames>().ToArray());

        using var other = Helpers.GetCsvImporter("FULL NAME,other name\na,b\n");
        Assert.Equal("b", other.ReadSheet().ReadRows<OrdinalNames>().Single().Name);
    }

    [Fact]
    public void ReadRows_SeveralColumnNameAttributesNoneFound_ThrowsWithFirstName()
    {
        using var importer = Helpers.GetCsvImporter("Other\na\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<SeveralNames>().ToArray());
        Assert.Contains("Can't find column \"Full Name\"", exception.Message);
        Assert.Equal("Full Name", exception.ColumnName);
        using var validating = Helpers.GetCsvImporter("Other\na\n");
        Assert.False(validating.ReadSheet().TryValidateHeading<SeveralNames>(out var problems));
        Assert.Contains("Full Name", Assert.Single(problems));
    }

    [Fact]
    public void ReadRows_SeveralColumnNameAttributesOptional_KeepsValue()
    {
        using var importer = Helpers.GetCsvImporter("Other\na\n");
        Assert.Equal("initial", importer.ReadSheet().ReadRows<OptionalSeveralNames>().Single().Name);
    }

    public class SeveralNames
    {
        [ExcelColumnName("Full Name")]
        [ExcelColumnName("Name")]
        public string? Name { get; set; }
    }

    public class OrdinalNames
    {
        [ExcelColumnName("Full Name", Comparison = StringComparison.Ordinal)]
        [ExcelColumnName("Other Name")]
        public string? Name { get; set; }
    }

    public class OptionalSeveralNames
    {
        [ExcelColumnName("Full Name")]
        [ExcelColumnName("Name")]
        [ExcelOptional]
        public string? Name { get; set; } = "initial";
    }

    [Fact]
    public void ReadRows_NoColumnNameMatches_ThrowsNoColumnsMatching()
    {
        using var importer = Helpers.GetCsvImporter("A,B\n1,2\n");
        importer.Configuration.RegisterClassMap<Single>(c => c.Map(r => r.Value).WithColumnMatching(n => n == "Missing"));
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Single>().ToArray());
        Assert.Contains("Can't find a matching column", exception.Message);
        using var validating = Helpers.GetCsvImporter("A,B\n1,2\n");
        validating.Configuration.RegisterClassMap<Single>(c => c.Map(r => r.Value).WithColumnMatching(n => n == "Missing"));
        Assert.False(validating.ReadSheet().TryValidateHeading<Single>(out var problems));
        Assert.Single(problems);
    }

    [Fact]
    public void ReadRows_NoColumnsMatch_ThrowsNoColumnsMatching()
    {
        using var importer = Helpers.GetCsvImporter("A,B\n1,2\n");
        importer.Configuration.RegisterClassMap<Many>(c => c.Map(r => r.Values).WithColumnsMatching(new Regex("^Q")));
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Many>().ToArray());
        Assert.Contains("Can't find a matching column", exception.Message);
    }

    [Fact]
    public void ReadRows_ColumnMatchingWithoutHeading_Throws()
    {
        using var importer = Helpers.GetCsvImporter("1,2\n");
        importer.Configuration.RegisterClassMap<Single>(c => c.Map(r => r.Value).WithColumnMatching(new Regex("A")));
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Single>().ToArray());
        Assert.Contains("There's no heading to find columns by name", exception.Message);

        using var other = Helpers.GetCsvImporter("1,2\n");
        other.Configuration.RegisterClassMap<Single>(c => c.Map(r => r.Value).WithColumnMatching(n => n == "A"));
        var otherSheet = other.ReadSheet();
        otherSheet.HasHeading = false;
        Assert.Throws<ExcelMappingException>(() => otherSheet.ReadRows<Single>().ToArray());
    }

    [Fact]
    public void ReadRows_SeveralColumnIndexAttributesNoneFound_Throws()
    {
        using var importer = Helpers.GetCsvImporter("A,B\n1,2\n");
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<SeveralIndices>().ToArray());
        using var optional = Helpers.GetCsvImporter("A,B\n1,2\n");
        Assert.Equal("initial", optional.ReadSheet().ReadRows<OptionalSeveralIndices>().Single().Value);
    }

    public class SeveralIndices
    {
        [ExcelColumnIndex(5)]
        [ExcelColumnIndex(6)]
        public string? Value { get; set; }
    }

    public class OptionalSeveralIndices
    {
        [ExcelColumnIndex(5)]
        [ExcelColumnIndex(6)]
        [ExcelOptional]
        public string? Value { get; set; } = "initial";
    }

    [Fact]
    public void ReadRows_RowShorterThanHeading_MissingCellsAreEmpty()
    {
        // A CSV row with fewer values than the heading: its missing cells are empty, not missing columns.
        using var importer = Helpers.GetCsvImporter("A,B\nx,1|2\ny\n");
        importer.Configuration.RegisterClassMap<ShortRow>(c =>
        {
            c.Map(r => r.A);
            c.Map(r => r.B).WithColumnIndex(1);
            c.Map(r => r.Values).WithColumnIndex(1).WithSeparators('|');
        });
        var rows = importer.ReadSheet().ReadRows<ShortRow>().ToArray();
        Assert.Equal("1|2", rows[0].B);
        Assert.Equal(["1", "2"], rows[0].Values);
        Assert.Null(rows[1].B);
        Assert.Empty(rows[1].Values);
    }

    public class Single
    {
        public string? Value { get; set; }
    }

    public class Many
    {
        public string[] Values { get; set; } = [];
    }

    public class ShortRow
    {
        public string? A { get; set; }
        public string? B { get; set; } = "initial";
        public string[] Values { get; set; } = ["initial"];
    }

    [Theory]
    [InlineData("\"Name \",Other\nAlice,x\n")]
    [InlineData("\" Name\",Other\nAlice,x\n")]
    [InlineData("\"  Name  \",Other\nAlice,x\n")]
    public void ReadRows_HeadingNameWithOuterWhiteSpace_FindsColumn(string csv)
    {
        using var importer = Helpers.GetCsvImporter(csv);
        var sheet = importer.ReadSheet();
        sheet.ValidateHeading<NamedRow>();
        Assert.Equal("Alice", sheet.ReadRows<NamedRow>().Single().Name);
    }

    [Fact]
    public void ReadRows_ExactAndPaddedHeadingNames_PrefersExactName()
    {
        using var importer = Helpers.GetCsvImporter("\"Name \",Name\nPadded,Exact\n");
        Assert.Equal("Exact", importer.ReadSheet().ReadRows<NamedRow>().Single().Name);
    }

    [Fact]
    public void ReadRows_TwoPaddedHeadingNames_ThrowsAmbiguousColumn()
    {
        using var importer = Helpers.GetCsvImporter("\"Name \",\" Name\"\na,b\n");
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<NamedRow>().ToArray());
        Assert.StartsWith("More than one column matches \"Name\" (columns A and B)", exception.Message);
    }

    [Fact]
    public void ReadRows_PaddedHeadingNameWithOrdinalComparison_MatchesCase()
    {
        using var importer = Helpers.GetCsvImporter("\"name \"\nAlice\n");
        importer.Configuration.RegisterClassMap<NamedRow>(c => c.Map(r => r.Name).WithColumnName("Name", StringComparison.Ordinal));
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<NamedRow>().ToArray());
    }

    public class NamedRow
    {
        public string? Name { get; set; }
    }
}
