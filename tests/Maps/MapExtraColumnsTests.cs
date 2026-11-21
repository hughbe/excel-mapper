using System.Dynamic;
using System.Linq;

namespace ExcelMapper.Tests;

public class MapExtraColumnsTests
{
    [Fact]
    public void ReadRows_AutoMappedExtraColumns_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,City,Age,Country\na,London,1,UK\nb,Paris,2,FR\n");
        var rows = importer.ReadSheet().ReadRows<ExtraColumnsRow>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
        Assert.Equal([1, 2], rows.Select(r => r.Age));
        Assert.Equal(new Dictionary<string, string?> { ["City"] = "London", ["Country"] = "UK" }, rows[0].Extra);
        Assert.Equal(new Dictionary<string, string?> { ["City"] = "Paris", ["Country"] = "FR" }, rows[1].Extra);
        Assert.Equal(["City", "Country"], rows[0].Extra.Keys);
    }

    public class ExtraColumnsRow
    {
        // Declared first: the columns of members declared after it are still excluded.
        [ExcelExtraColumns]
        public Dictionary<string, string?> Extra { get; set; } = [];

        public string? Name { get; set; }

        public int Age { get; set; }
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsNone_ReturnsEmpty()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\na,1\n");
        var row = importer.ReadSheet().ReadRows<ExtraColumnsRow>().Single();
        Assert.Empty(row.Extra);
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsSkipsEmptyColumnNames_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,,Age,City\na,x,1,London\n");
        var row = importer.ReadSheet().ReadRows<ExtraColumnsRow>().Single();
        Assert.Equal(new Dictionary<string, string?> { ["City"] = "London" }, row.Extra);
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsEmptyValue_ReturnsNull()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,City\na,1,\n");
        var row = importer.ReadSheet().ReadRows<ExtraColumnsRow>().Single();
        Assert.Equal(new Dictionary<string, string?> { ["City"] = null }, row.Extra);
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsObjectXlsx_ReturnsTypedValues()
    {
        var date = new DateTime(2024, 3, 4);
        using var importer = XlsxBuilder.Create(["Name", "Score", "Joined", "Active"], ["a", 2.5, date, true]);
        var row = importer.ReadSheet().ReadRows<ObjectExtraColumnsRow>().Single();
        Assert.Equal("a", row.Name);
        Assert.Equal(3, row.Extra.Count);
        Assert.Equal(2.5, row.Extra["Score"]);
        Assert.Equal(date, row.Extra["Joined"]);
        Assert.Equal(true, row.Extra["Active"]);
    }

    public class ObjectExtraColumnsRow
    {
        public string? Name { get; set; }

        [ExcelExtraColumns]
        public IReadOnlyDictionary<string, object?> Extra { get; set; } = new Dictionary<string, object?>();
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsTypedValues_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,Q1,Q2\na,1,2\n");
        var row = importer.ReadSheet().ReadRows<IntExtraColumnsRow>().Single();
        Assert.Equal(new Dictionary<string, int> { ["Q1"] = 1, ["Q2"] = 2 }, row.Extra);
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsInvalidTypedValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Name,Q1,Q2\na,1,x\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<IntExtraColumnsRow>().ToArray());
        Assert.Equal("C2", exception.CellAddress);
    }

    public class IntExtraColumnsRow
    {
        public string? Name { get; set; }

        [ExcelExtraColumns]
        public IDictionary<string, int> Extra { get; set; } = new Dictionary<string, int>();
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsExpandoObject_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,City\na,London\n");
        var row = importer.ReadSheet().ReadRows<ExpandoExtraColumnsRow>().Single();
        var extra = (IDictionary<string, object?>)row.Extra;
        Assert.Equal("London", Assert.Single(extra).Value);
        Assert.Equal("London", ((dynamic)row.Extra).City);
    }

    public class ExpandoExtraColumnsRow
    {
        public string? Name { get; set; }

        [ExcelExtraColumns]
        public ExpandoObject Extra { get; set; } = new();
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsWithColumnAttributes_ExcludesThoseColumns()
    {
        using var importer = Helpers.GetCsvImporter("Full Name,Code,Tags,Note,Other\na,1,x;y,n,o\n");
        var row = importer.ReadSheet().ReadRows<AttributedRow>().Single();
        Assert.Equal("a", row.Name);
        Assert.Equal("1", row.Code);
        Assert.Equal(["x", "y"], row.Tags);
        Assert.Equal(new Dictionary<string, string?> { ["Note"] = "n", ["Other"] = "o" }, row.Extra);
    }

    public class AttributedRow
    {
        [ExcelColumnName("Full Name")]
        public string? Name { get; set; }

        [ExcelColumnIndex(1)]
        public string? Code { get; set; }

        [ExcelSeparators(';')]
        public string[] Tags { get; set; } = [];

        [ExcelOptional]
        public string? Missing { get; set; }

        [ExcelRowNumber]
        public int Row { get; set; }

        [ExcelExtraColumns]
        public Dictionary<string, string?> Extra { get; set; } = [];
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsIgnoredMember_IncludesItsColumn()
    {
        using var importer = Helpers.GetCsvImporter("Name,Secret\na,s\n");
        var row = importer.ReadSheet().ReadRows<IgnoredMemberRow>().Single();
        Assert.Null(row.Secret);
        Assert.Equal(new Dictionary<string, string?> { ["Secret"] = "s" }, row.Extra);
    }

    public class IgnoredMemberRow
    {
        public string? Name { get; set; }

        [ExcelIgnore]
        public string? Secret { get; set; }

        [ExcelExtraColumns]
        public Dictionary<string, string?> Extra { get; set; } = [];
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsNestedObject_ExcludesNestedColumns()
    {
        using var importer = Helpers.GetCsvImporter("Name,Street,Town,Phone\na,High St,Leeds,123\n");
        var row = importer.ReadSheet().ReadRows<NestedRow>().Single();
        Assert.Equal("High St", row.Address.Street);
        Assert.Equal("Leeds", row.Address.Town);
        Assert.Equal(new Dictionary<string, string?> { ["Phone"] = "123" }, row.Extra);
    }

    public class NestedRow
    {
        public string? Name { get; set; }

        public Address Address { get; set; } = new();

        [ExcelExtraColumns]
        public Dictionary<string, string?> Extra { get; set; } = [];
    }

    public class Address
    {
        public string? Street { get; set; }
        public string? Town { get; set; }
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsIgnoreSeparatorsInColumnNames_ExcludesMatchedColumns()
    {
        using var importer = Helpers.GetCsvImporter("First Name,Last_Name,City\na,b,London\n");
        importer.Configuration.IgnoreSeparatorsInColumnNames = true;
        var row = importer.ReadSheet().ReadRows<SeparatorsRow>().Single();
        Assert.Equal("a", row.FirstName);
        Assert.Equal("b", row.LastName);
        Assert.Equal(new Dictionary<string, string?> { ["City"] = "London" }, row.Extra);
    }

    public class SeparatorsRow
    {
        public string? FirstName { get; set; }
        public string? LastName { get; set; }

        [ExcelExtraColumns]
        public Dictionary<string, string?> Extra { get; set; } = [];
    }

    [Fact]
    public void ReadRows_AutoMappedExtraColumnsRecord_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,City\na,London\n");
        var row = importer.ReadSheet().ReadRows<ExtraColumnsRecord>().Single();
        Assert.Equal("a", row.Name);
        Assert.Equal(new Dictionary<string, string> { ["City"] = "London" }, row.Extra);
    }

    public record ExtraColumnsRecord(string Name, [property: ExcelExtraColumns] Dictionary<string, string> Extra);

    [Fact]
    public void ReadRows_CustomMappedExtraColumns_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Full Name,City,Country\na,London,UK\n");
        importer.Configuration.RegisterClassMap<FluentRow>(c =>
        {
            c.Map(p => p.Name).WithColumnName("Full Name");
            c.MapExtraColumns(p => p.Extra);
        });
        var row = importer.ReadSheet().ReadRows<FluentRow>().Single();
        Assert.Equal("a", row.Name);
        Assert.Equal(new Dictionary<string, string> { ["City"] = "London", ["Country"] = "UK" }, row.Extra);
    }

    [Fact]
    public void ReadRows_CustomMappedExtraColumnsWithValueMap_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Full Name,City\na,london\n");
        importer.Configuration.RegisterClassMap<FluentRow>(c =>
        {
            c.Map(p => p.Name).WithColumnName("Full Name");
            c.MapExtraColumns(p => p.Extra).WithValueMap(v => v.WithConverter(s => s!.ToUpperInvariant()));
        });
        var row = importer.ReadSheet().ReadRows<FluentRow>().Single();
        Assert.Equal(new Dictionary<string, string> { ["City"] = "LONDON" }, row.Extra);
    }

    [Fact]
    public void ReadRows_CustomMappedExtraColumnsUnmappedMembers_IncludesTheirColumns()
    {
        // Only members in the class map read columns.
        using var importer = Helpers.GetCsvImporter("Name,City\na,London\n");
        importer.Configuration.RegisterClassMap<FluentRow>(c => c.MapExtraColumns(p => p.Extra));
        var row = importer.ReadSheet().ReadRows<FluentRow>().Single();
        Assert.Null(row.Name);
        Assert.Equal(new Dictionary<string, string> { ["Name"] = "a", ["City"] = "London" }, row.Extra);
    }

    [Fact]
    public void ReadRows_CustomMappedAttributedMember_HonoursAttribute()
    {
        using var importer = Helpers.GetCsvImporter("Name,City,Age\na,London,1\n");
        importer.Configuration.RegisterClassMap<ExtraColumnsRow>(c =>
        {
            c.Map(p => p.Name);
            c.Map(p => p.Extra);
            c.Map(p => p.Age);
        });
        var row = importer.ReadSheet().ReadRows<ExtraColumnsRow>().Single();
        Assert.Equal(new Dictionary<string, string?> { ["City"] = "London" }, row.Extra);
    }

    public class FluentRow
    {
        public string? Name { get; set; }
        public Dictionary<string, string> Extra { get; set; } = [];
    }

    [Fact]
    public void ReadRowResults_ExtraColumns_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,City,Age\na,London,1\nb,Paris,x\n");
        var results = importer.ReadSheet().ReadRowResults<ExtraColumnsRow>().ToArray();
        Assert.True(results[0].IsValid);
        Assert.Equal(new Dictionary<string, string?> { ["City"] = "London" }, results[0].Value!.Extra);
        Assert.Equal("C3", Assert.Single(results[1].Errors).CellAddress);
    }

    [Fact]
    public void ValidateHeading_ExtraColumns_Success()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,City\na,1,London\n");
        var sheet = importer.ReadSheet();
        sheet.ValidateHeading<ExtraColumnsRow>();
        Assert.Equal("London", sheet.ReadRows<ExtraColumnsRow>().Single().Extra["City"]);
    }

    [Fact]
    public void ReadRows_ExtraColumnsNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("a,London\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<NoHeadingRow>().ToArray());
    }

    public class NoHeadingRow
    {
        [ExcelColumnIndex(0)]
        public string? Name { get; set; }

        [ExcelExtraColumns]
        public Dictionary<string, string> Extra { get; set; } = [];
    }

    [Fact]
    public void ReadRows_ExtraColumnsNotDictionary_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Name,City\na,London\n");
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringExtraColumnsRow>().ToArray());
    }

    public class StringExtraColumnsRow
    {
        [ExcelExtraColumns]
        public string? Extra { get; set; }
    }

    [Fact]
    public void ReadRows_ExtraColumnsNonStringKeys_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Name,City\na,London\n");
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<IntKeyExtraColumnsRow>().ToArray());
    }

    public class IntKeyExtraColumnsRow
    {
        [ExcelExtraColumns]
        public Dictionary<int, string> Extra { get; set; } = [];
    }
}
