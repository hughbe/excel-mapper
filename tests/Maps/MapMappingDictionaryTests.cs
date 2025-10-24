using ExcelMapper.Mappers;

namespace ExcelMapper.Tests;

public class MapMappingDictionaryTests
{
    [Fact]
    public void ReadRow_AutoMappedMappingDictionaryAttribute_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MappingDictionaryClass>();
        Assert.Equal("a", row1.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row1.EnumValue);

        var row2 = sheet.ReadRow<MappingDictionaryClass>();
        Assert.Equal("extra", row2.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row2.EnumValue);

        var row3 = sheet.ReadRow<MappingDictionaryClass>();
        Assert.Equal("B", row3.StringValue);
        Assert.Equal(MapUsingValueEnum.Second, row3.EnumValue);

        var row4 = sheet.ReadRow<MappingDictionaryClass>();
        Assert.Null(row4.StringValue);
        Assert.Equal(MapUsingValueEnum.Unknown, row4.EnumValue);
    }

    private class MappingDictionaryClass
    {
        [ExcelMappingDictionary("b", "extra")]
        [ExcelInvalidValue("Invalid")]
        public string StringValue { get; set; } = default!;

        [ExcelMappingDictionary("one", MapUsingValueEnum.First)]
        [ExcelInvalidValue(MapUsingValueEnum.Unknown)]
        public MapUsingValueEnum EnumValue { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedMappingDictionaryAttribute_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<MappingDictionaryClass>(c =>
        {
            c.Map(o => o.StringValue);
            c.Map(o => o.EnumValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MappingDictionaryClass>();
        Assert.Equal("a", row1.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row1.EnumValue);

        var row2 = sheet.ReadRow<MappingDictionaryClass>();
        Assert.Equal("extra", row2.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row2.EnumValue);

        var row3 = sheet.ReadRow<MappingDictionaryClass>();
        Assert.Equal("B", row3.StringValue);
        Assert.Equal(MapUsingValueEnum.Second, row3.EnumValue);

        var row4 = sheet.ReadRow<MappingDictionaryClass>();
        Assert.Null(row4.StringValue);
        Assert.Equal(MapUsingValueEnum.Unknown, row4.EnumValue);
    }

    [Fact]
    public void ReadRow_AutoMappedMappingDictionaryAttributeComparer_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MappingDictionaryComparerClass>();
        Assert.Equal("a", row1.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row1.EnumValue);

        var row2 = sheet.ReadRow<MappingDictionaryComparerClass>();
        Assert.Equal("extra", row2.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row2.EnumValue);

        var row3 = sheet.ReadRow<MappingDictionaryComparerClass>();
        Assert.Equal("extra", row3.StringValue);
        Assert.Equal(MapUsingValueEnum.Second, row3.EnumValue);

        var row4 = sheet.ReadRow<MappingDictionaryComparerClass>();
        Assert.Null(row4.StringValue);
        Assert.Equal(MapUsingValueEnum.Unknown, row4.EnumValue);
    }

    private class MappingDictionaryComparerClass
    {
        [ExcelMappingDictionary("b", "extra")]
        [ExcelMappingDictionaryComparer(StringComparison.InvariantCultureIgnoreCase)]
        [ExcelInvalidValue("Invalid")]
        public string StringValue { get; set; } = default!;

        [ExcelMappingDictionary("one", MapUsingValueEnum.First)]
        [ExcelInvalidValue(MapUsingValueEnum.Unknown)]
        public MapUsingValueEnum EnumValue { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedMappingDictionaryAttributeComparer_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<MappingDictionaryComparerClass>(c =>
        {
            c.Map(o => o.StringValue);
            c.Map(o => o.EnumValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MappingDictionaryComparerClass>();
        Assert.Equal("a", row1.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row1.EnumValue);

        var row2 = sheet.ReadRow<MappingDictionaryComparerClass>();
        Assert.Equal("extra", row2.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row2.EnumValue);

        var row3 = sheet.ReadRow<MappingDictionaryComparerClass>();
        Assert.Equal("extra", row3.StringValue);
        Assert.Equal(MapUsingValueEnum.Second, row3.EnumValue);

        var row4 = sheet.ReadRow<MappingDictionaryComparerClass>();
        Assert.Null(row4.StringValue);
        Assert.Equal(MapUsingValueEnum.Unknown, row4.EnumValue);
    }

    [Fact]
    public void ReadRow_AutoMappedMappingDictionaryAttributeComparerRequired_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MappingDictionaryComparerRequiredClass>();
        Assert.Equal("Invalid", row1.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row1.EnumValue);

        var row2 = sheet.ReadRow<MappingDictionaryComparerRequiredClass>();
        Assert.Equal("extra", row2.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row2.EnumValue);

        var row3 = sheet.ReadRow<MappingDictionaryComparerRequiredClass>();
        Assert.Equal("extra", row3.StringValue);
        Assert.Equal(MapUsingValueEnum.Second, row3.EnumValue);

        var row4 = sheet.ReadRow<MappingDictionaryComparerRequiredClass>();
        Assert.Null(row4.StringValue);
        Assert.Equal(MapUsingValueEnum.Unknown, row4.EnumValue);
    }

    private class MappingDictionaryComparerRequiredClass
    {
        [ExcelMappingDictionary("b", "extra")]
        [ExcelMappingDictionaryComparer(StringComparison.InvariantCultureIgnoreCase)]
        [ExcelMappingDictionaryBehavior(MappingDictionaryMapperBehavior.Required)]
        [ExcelInvalidValue("Invalid")]
        public string StringValue { get; set; } = default!;

        [ExcelMappingDictionary("one", MapUsingValueEnum.First)]
        [ExcelInvalidValue(MapUsingValueEnum.Unknown)]
        public MapUsingValueEnum EnumValue { get; set; }
    }

    [Fact]
    public void ReadRow_WithMappingMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<WithMappingValue>(c =>
        {
            c.Map(o => o.StringValue)
                .WithMapping(new Dictionary<string, string>
                {
                    { "b", "extra" }
                })
                .WithInvalidFallback("Missing");

            c.Map(o => o.EnumValue)
                .WithMapping(new Dictionary<string, MapUsingValueEnum>
                {
                    { "one", MapUsingValueEnum.First }
                })
                .WithInvalidFallback(MapUsingValueEnum.Unknown);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<WithMappingValue>();
        Assert.Equal("a", row1.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row1.EnumValue);

        var row2 = sheet.ReadRow<WithMappingValue>();
        Assert.Equal("extra", row2.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row2.EnumValue);

        var row3 = sheet.ReadRow<WithMappingValue>();
        Assert.Equal("B", row3.StringValue);
        Assert.Equal(MapUsingValueEnum.Second, row3.EnumValue);

        var row4 = sheet.ReadRow<WithMappingValue>();
        Assert.Null(row4.StringValue);
        Assert.Equal(MapUsingValueEnum.Unknown, row4.EnumValue);
    }

    [Fact]
    public void ReadRow_WithMappingMapComparer_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<WithMappingValue>(c =>
        {
            c.Map(o => o.StringValue)
                .WithMapping(new Dictionary<string, string>
                {
                    { "b", "extra" }
                }, StringComparer.OrdinalIgnoreCase)
                .WithInvalidFallback("Missing");

            c.Map(o => o.EnumValue)
                .WithMapping(new Dictionary<string, MapUsingValueEnum>
                {
                    { "one", MapUsingValueEnum.First }
                })
                .WithInvalidFallback(MapUsingValueEnum.Unknown);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<WithMappingValue>();
        Assert.Equal("a", row1.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row1.EnumValue);

        var row2 = sheet.ReadRow<WithMappingValue>();
        Assert.Equal("extra", row2.StringValue);
        Assert.Equal(MapUsingValueEnum.First, row2.EnumValue);

        var row3 = sheet.ReadRow<WithMappingValue>();
        Assert.Equal("extra", row3.StringValue);
        Assert.Equal(MapUsingValueEnum.Second, row3.EnumValue);

        var row4 = sheet.ReadRow<WithMappingValue>();
        Assert.Null(row4.StringValue);
        Assert.Equal(MapUsingValueEnum.Unknown, row4.EnumValue);
    }

    private enum MapUsingValueEnum
    {
        First,
        Second,
        Unknown
    }

    private class WithMappingValue
    {
        public string StringValue { get; set; } = default!;
        public MapUsingValueEnum EnumValue { get; set; }
    }

    [Fact]
    public void ReadRow_WithRequiredMappingMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<WithMappingValue>(c =>
        {
            c.Map(o => o.StringValue)
                .WithMapping(new Dictionary<string, string>
                    {
                        { "a", "12345" }
                    }, behavior: MappingDictionaryMapperBehavior.Required)
                .WithColumnIndex(0)
                .WithInvalidFallback("Missing");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<WithMappingValue>();
        Assert.Equal("12345", row1.StringValue);

        var row2 = sheet.ReadRow<WithMappingValue>();
        Assert.Equal("Missing", row2.StringValue);

        var row3 = sheet.ReadRow<WithMappingValue>();
        Assert.Equal("Missing", row3.StringValue);

        var row4 = sheet.ReadRow<WithMappingValue>();
        Assert.Null(row4.StringValue);
    }
}
