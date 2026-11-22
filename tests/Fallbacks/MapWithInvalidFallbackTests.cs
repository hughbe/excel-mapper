namespace ExcelMapper.Tests;

public class MapWithInvalidFallbackTests
{
    [Fact]
    public void ReadRow_CustomMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value)
                .WithInvalidFallback("empty");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("  value  ", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Null(row3.Value);
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
                .WithInvalidFallback(null);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("  value  ", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Null(row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedInt_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<IntValue>(c =>
        {
            c.Map(o => o.Value)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<IntValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntValue>());

        // Invalid cell value.
        var row3 = sheet.ReadRow<IntValue>();
        Assert.Equal(10, row3.Value);
    }

    private class IntValue
    {
        public int Value { get; set; } = default!;
    }

    [Fact]
    public void Map_IntNull_ThrowsArgumentNullException()
    {
        // A non-nullable member can't have a null value, so null is a null ICellFallback<int>.
        var map = new ExcelClassMap<IntValue>();
        Assert.Throws<ArgumentNullException>("fallback", () => map.Map(o => o.Value).WithInvalidFallback(null!));
    }
}
