namespace ExcelMapper.Tests;

public class StringSplitValueTests
{
    [Fact]
    public void ReadRow_AutoMappedSingleCharSeparatorsAttributeArrayMap_ReturnsExpected()
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
        [ExcelSeparators(",")]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSingleCharSeparatorsAttributeArrayMap_ReturnsExpected()
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
    public void ReadRow_AutoMappedSingleCharSeparatorsAttributeArrayMapProcess_ReturnsExpected()
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
    public void ReadRow_DefaultMappedSingleCharSeparatorsAttributeArrayMapProcess_ReturnsExpected()
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
    public void ReadRow_AutoMappedSingleRemoveEmptyEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
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
        [ExcelSeparators(",", Options = StringSplitOptions.RemoveEmptyEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSingleRemoveEmptyEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
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
    public void ReadRow_AutoMappedSingleTrimEntriesSeparatorsAttributeArrayMapProcess_ReturnsExpected()
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
        [ExcelSeparators(",", Options = StringSplitOptions.TrimEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSingleTrimEntriesSeparatorsAttributeArrayMapProcess_ReturnsExpected()
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
    public void ReadRow_AutoMappedSingleTrimEntriesSeparatorsAttributeArrayMapNoProcess_ReturnsExpected()
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
    public void ReadRow_DefaultMappedSingleTrimEntriesSeparatorsAttributeArrayMapNoProcess_ReturnsExpected()
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
    public void ReadRow_AutoMappedSingleRemoveEmptyEntriesTrimEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
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
        [ExcelSeparators(",", Options = StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSingleRemoveEmptyEntriesTrimEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
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
    public void ReadRow_AutoMappedSingleRemoveEmptyEntriesTrimEntriesSeparatorsAttributeArrayMapNoProcess_ReturnsExpected()
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
    public void ReadRow_DefaultMappedSingleRemoveEmptyEntriesTrimEntriesSeparatorsAttributeArrayMapNoProcess_ReturnsExpected()
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
    public void ReadRow_AutoMappedSingleStringSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleStringSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleStringSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleStringSeparatorsClass>();
        Assert.Equal(new string[] { "  1  " }, row3.Value);

        var row4 = sheet.ReadRow<SingleStringSeparatorsClass>();
        Assert.Equal(new string[] { "  " }, row4.Value);
    }

    private class SingleStringSeparatorsClass
    {
        [ExcelSeparators(", ")]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSingleStringSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");
        importer.Configuration.RegisterClassMap<SingleStringSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleStringSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleStringSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleStringSeparatorsClass>();
        Assert.Equal(new string[] { "  1  " }, row3.Value);

        var row4 = sheet.ReadRow<SingleStringSeparatorsClass>();
        Assert.Equal(new string[] { "  " }, row4.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedSingleStringTrimEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleStringTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleStringTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleStringTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleStringTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { null }, row4.Value);
    }

    private class SingleStringTrimEntriesSeparatorsClass
    {
        [ExcelSeparators(", ", Options = StringSplitOptions.TrimEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSingleStringTrimEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");
        importer.Configuration.RegisterClassMap<SingleStringTrimEntriesSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleStringTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleStringTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleStringTrimEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleStringTrimEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { null }, row4.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedSingleStringRemoveEmptyEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleStringRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleStringRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleStringRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "  1  " }, row3.Value);

        var row4 = sheet.ReadRow<SingleStringRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "  " }, row4.Value);
    }

    private class SingleStringRemoveEmptyEntriesSeparatorsClass
    {
        [ExcelSeparators(", ", Options = StringSplitOptions.RemoveEmptyEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSingleStringRemoveEmptyEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");
        importer.Configuration.RegisterClassMap<SingleStringRemoveEmptyEntriesSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleStringRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleStringRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleStringRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "  1  " }, row3.Value);

        var row4 = sheet.ReadRow<SingleStringRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "  " }, row4.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedSingleStringTrimEntriesRemoveEmptyEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleStringTrimEntriesRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleStringTrimEntriesRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleStringTrimEntriesRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleStringTrimEntriesRemoveEmptyEntriesSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    private class SingleStringTrimEntriesRemoveEmptyEntriesSeparatorsClass
    {
        [ExcelSeparators(", ", Options = StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSingleStringTrimEntriesRemoveEmptyEntriesSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithComma_Space.xlsx");
        importer.Configuration.RegisterClassMap<SingleStringTrimEntriesRemoveEmptyEntriesSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SingleStringTrimEntriesRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

        var row2 = sheet.ReadRow<SingleStringTrimEntriesRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string?[] { "1", "2" }, row2.Value);

        var row3 = sheet.ReadRow<SingleStringTrimEntriesRemoveEmptyEntriesSeparatorsClass>();
        Assert.Equal(new string[] { "1" }, row3.Value);

        var row4 = sheet.ReadRow<SingleStringTrimEntriesRemoveEmptyEntriesSeparatorsClass>();
        Assert.Empty(row4.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MultipleSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

        var row2 = sheet.ReadRow<MultipleSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "3" }, row2.Value);

        var row3 = sheet.ReadRow<MultipleSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3" }, row3.Value);
    }

    private class MultipleSeparatorsClass
    {
        [ExcelSeparators(";", ",")]
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedSeparatorsAttributeArrayMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
        importer.Configuration.RegisterClassMap<MultipleSeparatorsClass>(c =>
        {
            c.Map(p => p.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MultipleSeparatorsClass>();
        Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

        var row2 = sheet.ReadRow<MultipleSeparatorsClass>();
        Assert.Equal(new string?[] { "1", null, "3" }, row2.Value);

        var row3 = sheet.ReadRow<MultipleSeparatorsClass>();
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
