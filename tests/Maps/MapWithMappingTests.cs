using System;
using System.Collections.Generic;
using ExcelMapper.Mappers;
using Xunit;

namespace ExcelMapper.Tests;

public class MapWithMappingTests
{
    [Fact]
    public void ReadRow_WithMappingMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<WithMappingValueMap>();

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
    [Fact]
    public void ReadRow_WithRequiredMappingMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<WithMappingValueRequiredMap>();

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

    private class WithMappingValueMap : ExcelClassMap<WithMappingValue>
    {
        public WithMappingValueMap()
        {
            Map(c => c.StringValue)
                .WithMapping(new Dictionary<string, string>
                {
                    { "b", "extra" }
                }, StringComparer.OrdinalIgnoreCase);

            Map(c => c.EnumValue)
                .WithMapping(new Dictionary<string, MapUsingValueEnum>
                {
                    { "one", MapUsingValueEnum.First }
                })
                .WithInvalidFallback(MapUsingValueEnum.Unknown);
        }
    }

    private class WithMappingValueRequiredMap : ExcelClassMap<WithMappingValue>
    {
        public WithMappingValueRequiredMap()
        {
            Map(c => c.StringValue)
                .WithMapping(new Dictionary<string, string>
                    {
                        { "a", "12345" }
                    }, behavior: DictionaryMapperBehavior.Required)
                .WithColumnIndex(0)
                .WithInvalidFallback("Missing");
        }
    }
}
