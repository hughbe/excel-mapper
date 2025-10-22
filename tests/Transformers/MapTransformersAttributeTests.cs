using ExcelMapper.Abstractions;
using ExcelMapper.Transformers;

namespace ExcelMapper.Tests;

public class MapTransformersAttributeTests
{
    [Fact]
    public void ReadRow_AutoMappedWithTransformer_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("VALUE", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("  VALUE  ", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Null(row3.Value);
    }

    private class StringValue
    {
        [ExcelTransformer(typeof(UpperStringCellTransformer), ConstructorArguments = new object[] { false })]
        public string Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedWithTransformer_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("VALUE", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("  VALUE  ", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Null(row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedWithEmptyTransformers_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value)
                .WithTransformers([]);
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

    private class MultipleStringValue
    {
        [ExcelTransformer(typeof(TrimStringCellTransformer))]
        [ExcelTransformer(typeof(UpperStringCellTransformer), ConstructorArguments = new object[] { true })]
        public string Value { get; set; } = default!;
    }

    private class UpperStringCellTransformer : ICellTransformer
    {
        public bool _isTrimmed;

        public UpperStringCellTransformer(bool isTrimmed)
        {
            _isTrimmed = isTrimmed;
        }

        public string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
        {
            var value = readResult.GetString();
            if (_isTrimmed)
            {
                Assert.Equal(value, value?.Trim());
            }
            return value?.ToUpperInvariant();
        }
    }
}
