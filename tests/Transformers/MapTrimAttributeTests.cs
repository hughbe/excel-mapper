namespace ExcelMapper.Tests;

public class MapTrimAttributeTests
{
    [Fact]
    public void ReadRow_AutoMappedTrimAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringTrimValue>();
        Assert.Equal("value", row1.Value);

        var row2 = sheet.ReadRow<StringTrimValue>();
        Assert.Equal("value", row2.Value);

        var row3 = sheet.ReadRow<StringTrimValue>();
        Assert.Null(row3.Value);
    }

    private class StringTrimValue
    {
        [ExcelTrimString]
        public string Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedTrimAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringTrimValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringTrimValue>();
        Assert.Equal("value", row1.Value);

        var row2 = sheet.ReadRow<StringTrimValue>();
        Assert.Equal("value", row2.Value);

        var row3 = sheet.ReadRow<StringTrimValue>();
        Assert.Null(row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedTrimAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringTrimValue>(c =>
        {
            c.Map(o => o.Value)
                .WithTransformers([]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringTrimValue>();
        Assert.Equal("value", row1.Value);

        var row2 = sheet.ReadRow<StringTrimValue>();
        Assert.Equal("  value  ", row2.Value);

        var row3 = sheet.ReadRow<StringTrimValue>();
        Assert.Null(row3.Value);
    }
}
