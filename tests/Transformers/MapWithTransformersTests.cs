
namespace ExcelMapper.Tests;

public class MapWithTransformersTests
{
    [Fact]
    public void ReadRow_CustomMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value)
                .WithTransformers(new UpperStringCellTransformer(false));
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
    public void ReadRow_MultipleCustomMappedString_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(o => o.Value)
                .WithTransformers([new TrimTransformer(), new UpperStringCellTransformer(true)]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("VALUE", row1.Value);

        var row2 = sheet.ReadRow<StringValue>();
        Assert.Equal("VALUE", row2.Value);

        var row3 = sheet.ReadRow<StringValue>();
        Assert.Null(row3.Value);
    }

    private class StringValue
    {
        public string Value { get; set; } = default!;
    }

    private class UpperStringCellTransformer : ICellTransformer
    {
        public bool _isTrimmed;

        public UpperStringCellTransformer(bool isTrimmed)
        {
            _isTrimmed = isTrimmed;
        }

        public CellValue Transform(CellValue value, CellContext context)
        {
            var text = value.Text;
            if (_isTrimmed)
            {
                Assert.Equal(text, text?.Trim());
            }
            return new CellValue(text?.ToUpperInvariant());
        }
    }

    private class TrimTransformer : ICellTransformer
    {
        public CellValue Transform(CellValue value, CellContext context) => new CellValue(value.Text?.Trim());
    }
}
