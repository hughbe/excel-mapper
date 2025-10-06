using Xunit;

namespace ExcelMapper.Tests;

public class MapStringTests
{
    [Fact]
    public void ReadRow_String_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<string>();
        Assert.Equal("value", row1);

        // Valid value
        var row2 = sheet.ReadRow<string>();
        Assert.Equal("  value  ", row2);

        // Empty value
        var row3 = sheet.ReadRow<string>();
        Assert.Null(row3);

        // Last row.
        var row4 = sheet.ReadRow<string>();
        Assert.Equal("value", row4);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<string>());
    }

    [Fact]
    public void ReadRow_AutoMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<StringClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<StringClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<StringClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<StringClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<DefaultStringClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<StringClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<StringClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<StringClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<StringClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<CustomStringClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<StringClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<StringClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<StringClass>();
        Assert.Equal("empty", row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<StringClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringClass>());
    }

    private class StringClass
    {
        public string Value { get; set; } = default!;
    }

    private class DefaultStringClassMap : ExcelClassMap<StringClass>
    {
        public DefaultStringClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomStringClassMap : ExcelClassMap<StringClass>
    {
        public CustomStringClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback("empty")
                .WithInvalidFallback("invalid");
        }
    }
}
