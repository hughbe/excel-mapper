namespace ExcelMapper.Tests;

public class MapWithEmptyFallbackTests
{
    [Fact]
    public void ReadRow_CustomMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback("empty");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty cell value.
        var row3 = sheet.ReadRow<StringValue>();
        Assert.Equal("empty", row3.Value);
    }


    private class StringValue
    {
        public string Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_CustomMappedStringNull_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(null);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty cell value.
        var row3 = sheet.ReadRow<StringValue>();
        Assert.Null(row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedInvalid_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty cell value.
        Assert.Throws<InvalidCastException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedInt_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<IntValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<IntValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<IntValue>();
        Assert.Equal(10, row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntValue>());
    }

    private class IntValue
    {
        public int Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_CustomMappedIntNull_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<IntValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(null);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<IntValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<NullReferenceException>(() => sheet.ReadRow<IntValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedIntInvalid_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<IntValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback("fallback");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<IntValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<InvalidCastException>(() => sheet.ReadRow<IntValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntValue>());
    }
}
