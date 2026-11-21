namespace ExcelMapper.Tests;

public class MapWithTrimTests
{
    [Fact]
    public void ReadRow_CustomMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value)
                .WithTrim();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("value", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Null(row3.Value);
    }

    private class StringValue
    {
        public string Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DateCellWithFormats_KeepsDate()
    {
        // The cell is a real date. Trimming used to turn it into culture-formatted text, which then
        // failed to parse with the custom format.
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTime>(c => c.Map(p => p).WithFormats("yyyy-MM-dd").WithTrim());

        var sheet = importer.ReadSheet();
        Assert.Equal(new DateTime(2017, 07, 19), sheet.ReadRow<DateTime>());
    }
}
