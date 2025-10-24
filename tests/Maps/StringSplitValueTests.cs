namespace ExcelMapper.Tests;

public class StringSplitValueTests
{
    [Fact]
    public void ReadRow_AutoMappedSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

        var row2 = sheet.ReadRow<SeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "3" }, row2.Value);

        var row3 = sheet.ReadRow<SeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row3.Value);
    }

    private class SeparatorsClass
    {
        [ExcelSeparators(";", ",")]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
        importer.Configuration.RegisterClassMap<SeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

        var row2 = sheet.ReadRow<SeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "3" }, row2.Value);

        var row3 = sheet.ReadRow<SeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row3.Value);
    }
    [Fact]
    public void ReadRow_AutoMappedSeparatorsOptionsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SeparatorsOptionsClass>();
        Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

        var row2 = sheet.ReadRow<SeparatorsOptionsClass>();
        Assert.Equal(new string?[] { "1", "3" }, row2.Value);

        var row3 = sheet.ReadRow<SeparatorsOptionsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row3.Value);
    }

    private class SeparatorsOptionsClass
    {
        [ExcelSeparators(";", ",", Options = StringSplitOptions.RemoveEmptyEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSeparatorsOptionsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
        importer.Configuration.RegisterClassMap<SeparatorsOptionsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SeparatorsOptionsClass>();
        Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

        var row2 = sheet.ReadRow<SeparatorsOptionsClass>();
        Assert.Equal(new string?[] { "1", "3" }, row2.Value);

        var row3 = sheet.ReadRow<SeparatorsOptionsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row3.Value);
    }

    [Fact]
    public void ReadRow_SeparatorsArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
        importer.Configuration.RegisterClassMap<AutoSplitWithSeparatorClass>(c =>
        {
            c.Map(p => p.Value)
                .WithSeparators(";", ",");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
        Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

        var row2 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
        Assert.Equal(new string?[] { "1", null, "3" }, row2.Value);

        var row3 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row3.Value);
    }

    private class AutoSplitWithSeparatorClass
    {
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_IEnumerableSeparatorsMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
        importer.Configuration.RegisterClassMap<AutoSplitWithSeparatorClass>(c =>
        {
            c.Map(p => p.Value)
                .WithSeparators(new List<string> { ";", "," });
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
        Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

        var row2 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
        Assert.Equal(new string?[] { "1", null, "3" }, row2.Value);

        var row3 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row3.Value);
    }

    [Fact]
    public void ReadRow_MultiMapMissingRow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
        importer.Configuration.RegisterClassMap<MissingColumnRow>(c =>
        {
            c.Map(p => p.MissingValue)
                .WithSeparators(";", ",");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnRow>());
    }

    private class MissingColumnRow
    {
        public int[] MissingValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_MultiMapOptionalMissingRow_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
        importer.Configuration.RegisterClassMap<MissingColumnRow>(c =>
        {
            c.Map(p => p.MissingValue)
                .WithSeparators(";", ",")
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        MissingColumnRow row = sheet.ReadRow<MissingColumnRow>();
        Assert.Null(row.MissingValue);
    }
}
