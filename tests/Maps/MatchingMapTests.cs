using Xunit;

namespace ExcelMapper.Tests;

public class MatchingMapTests
{
    [Fact]
    public void ReadRow_ColumnNamesHasHeader_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<ColumnMatchingColumnNamesClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("1", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("2", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Equal("abc", row3.Value);
    }

    [Fact]
    public void ReadRow_ColumnNamesNoHeader_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
        importer.Configuration.RegisterClassMap<ColumnMatchingColumnNamesClassMap>();

        ExcelSheet sheet = importer.ReadSheet();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_ColumnNamesNothingMatchingNotOptional_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoColumnMatchingColumnNamesClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_ColumnNamesNothingMatchingOptional_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoColumnMatchingColumnNamesOptionalClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Null(row1.Value);
    }

    [Fact]
    public void ReadRow_PredicateHasHeader_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<ColumnMatchingPredicateClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("1", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("2", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Equal("abc", row3.Value);
    }

    [Fact]
    public void ReadRow_PredicateNoHeader_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
        importer.Configuration.RegisterClassMap<ColumnMatchingPredicateClassMap>();

        ExcelSheet sheet = importer.ReadSheet();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_PredicateNothingMatchingNotOptional_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoColumnMatchingPredicateClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_PredicateNothingMatchingOptional_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<NoColumnMatchingPredicateOptionalClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Null(row1.Value);
    }

    private class StringValue
    {
        public string Value { get; set; } = default!;
    }

    private class NoColumnMatchingColumnNamesClassMap : ExcelClassMap<StringValue>
    {
        public NoColumnMatchingColumnNamesClassMap()
        {
            Map(o => o.Value)
                .WithColumnNameMatching("NoSuchColumn");
        }
    }

    private class NoColumnMatchingColumnNamesOptionalClassMap : ExcelClassMap<StringValue>
    {
        public NoColumnMatchingColumnNamesOptionalClassMap()
        {
            Map(o => o.Value)
                .WithColumnNameMatching("NoSuchColumn")
                .MakeOptional();
        }
    }

    private class ColumnMatchingColumnNamesClassMap : ExcelClassMap<StringValue>
    {
        public ColumnMatchingColumnNamesClassMap()
        {
            Map(o => o.Value)
                .WithColumnNameMatching("NoSuchColumn", "Int Value");
        }
    }

    private class ColumnMatchingPredicateClassMap : ExcelClassMap<StringValue>
    {
        public ColumnMatchingPredicateClassMap()
        {
            int i = 0;

            Map(o => o.Value)
                .WithColumnNameMatching(s =>
                {
                    if (s == "Int Value" && i == 0)
                    {
                        i++;
                        return true;
                    }

                    if (s == "StringValue")
                    {
                        i++;
                        return true;
                    }

                    return false;
                });
        }
    }

    private class NoColumnMatchingPredicateClassMap : ExcelClassMap<StringValue>
    {
        public NoColumnMatchingPredicateClassMap()
        {
            Map(o => o.Value)
                .WithColumnNameMatching(_ => false);
        }
    }

    private class NoColumnMatchingPredicateOptionalClassMap : ExcelClassMap<StringValue>
    {
        public NoColumnMatchingPredicateOptionalClassMap()
        {
            Map(o => o.Value)
                .WithColumnNameMatching(_ => false)
                .MakeOptional();
        }
    }
}
