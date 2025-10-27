namespace ExcelMapper.Tests;

public class CharSplitValueTests
{
    [Fact]
    public void ReadRow_AutoMappedCharSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    private class SingleCharSeparatorsClass
    {
        [ExcelSeparators(',')]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedCharSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<SingleCharSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedCharSeparatorsAttributeArrayMapProcess_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string[] { "1", " 2", " 3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string?[] { "1", " ", " 2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string[] { "  1  " }, row3.Value);

        var row4 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(["  "], row4.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedCharSeparatorsAttributeArrayMapProcess_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");
        importer.Configuration.RegisterClassMap<SingleCharSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string[] { "1", " 2", " 3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string?[] { "1", " ", " 2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(new string[] { "  1  " }, row3.Value);

        var row4 = sheet.ReadRow<SingleCharSeparatorsClass>();
        Assert.Equal(["  "], row4.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedRemoveEmptyEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleRemoveEmptyEntriesSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    private class SingleRemoveEmptyEntriesSeparatorsClass
    {
        [ExcelSeparators(',', Options = StringSplitOptions.RemoveEmptyEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedRemoveEmptyEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<SingleCharSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleRemoveEmptyEntriesSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedTrimEntriesSeparatorsAttributeArrayMapProcess_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { null }, row4.Value);
    }

    private class SingleTrimEntriesSeparatorsClass
    {
        [ExcelSeparators(',', Options = StringSplitOptions.TrimEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedTrimEntriesSeparatorsAttributeArrayMapProcess_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");
        importer.Configuration.RegisterClassMap<SingleCharSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { null }, row4.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedTrimEntriesSeparatorsAttributeArrayMapNoProcess_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedTrimEntriesSeparatorsAttributeArrayMapNoProcess_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<SingleCharSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleTrimEntriesSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedRemoveEmptyEntriesTrimEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    private class SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass
    {
        [ExcelSeparators(',', Options = StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedRemoveEmptyEntriesTrimEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");
        importer.Configuration.RegisterClassMap<SingleCharSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedRemoveEmptyEntriesTrimEntriesSeparatorsAttributeArrayMapNoProcess_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedRemoveEmptyEntriesTrimEntriesSeparatorsAttributeArrayMapNoProcess_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
        importer.Configuration.RegisterClassMap<SingleCharSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleRemoveEmptyEntriesTrimEntriesSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

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
        [ExcelSeparators(';', ',')]
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
    public void ReadRow_AutoMappedSeparatorsAttributeOptionsArrayMap_ReturnsExpected()
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
        [ExcelSeparators(';', ',', Options = StringSplitOptions.RemoveEmptyEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSeparatorsAttributeOptionsArrayMap_ReturnsExpected()
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
                .WithSeparators(';', ',');
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
        Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

        var row2 = sheet.ReadRow<SeparatorsClass>();
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
                .WithSeparators(new List<char> { ';', ',' });
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
        Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

        var row2 = sheet.ReadRow<SeparatorsClass>();
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
                .WithSeparators(';', ',');
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
                .WithSeparators(';', ',')
                .MakeOptional();
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        MissingColumnRow row = sheet.ReadRow<MissingColumnRow>();
        Assert.Null(row.MissingValue);
    }
}
