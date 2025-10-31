namespace ExcelMapper.Tests;

public class MapStringTests
{
    [Fact]
    public void ReadRow_String_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
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
    public void ReadRow_DefaultMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<string>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
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
    public void ReadRow_CustomMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<string>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback("empty")
                .WithInvalidFallback("invalid");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<string>();
        Assert.Equal("value", row1);

        // Valid value
        var row2 = sheet.ReadRow<string>();
        Assert.Equal("  value  ", row2);

        // Empty value
        var row3 = sheet.ReadRow<string>();
        Assert.Equal("empty", row3);

        // Last row.
        var row4 = sheet.ReadRow<string>();
        Assert.Equal("value", row4);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<string>());
    }

    [Fact]
    public void ReadRow_AutoMappedStringValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<StringValue>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedStringValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<StringValue>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedStringValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback("empty")
                .WithInvalidFallback("invalid");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<StringValue>();
        Assert.Equal("empty", row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    private class StringValue
    {
        public string Value { get; set; } = default!;
    }
}
