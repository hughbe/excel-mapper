using System.Globalization;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// Transforming cells with an <see cref="ICellTransformer"/>, set with WithTransformers or
/// <see cref="ExcelTransformerAttribute"/>.
/// </summary>
public class MapCellTransformerTests
{
    [Fact]
    public void ReadRows_WithTransformers_TransformsInOrder()
    {
        using var importer = XlsxBuilder.Create(["Name"], ["  abc  "]);
        importer.Configuration.RegisterClassMap<StringRow>(c => c.Map(r => r.Name).WithTransformers(new TrimTransformer(), new UpperTransformer()));
        Assert.Equal("ABC", importer.ReadSheet().ReadRows<StringRow>().Single().Name);
    }

    [Fact]
    public void ReadRows_Attribute_Transforms()
    {
        using var importer = Helpers.GetCsvImporter("Name,Amount\nabc,\"1 234,5\"\n");
        var row = importer.ReadSheet().ReadRows<AttributeRow>().Single();
        Assert.Equal("ABC", row.Name);
        Assert.Equal(1234.5, row.Amount);
    }

    [Fact]
    public void ReadRows_TransformerReturnsNumber_ConvertsWithoutText()
    {
        // A transformer can return a typed value, which is converted as a number cell would be.
        using var importer = Helpers.GetCsvImporter("Amount\n\"1 234,5\"\n");
        importer.Configuration.RegisterClassMap<NumberRow>(c => c.Map(r => r.Amount).WithTransformers(new FrenchNumberTransformer()));
        Assert.Equal(1234.5, importer.ReadSheet().ReadRows<NumberRow>().Single().Amount);
    }

    [Fact]
    public void ReadRows_TransformerKeepsTypedCell_ReceivesNumber()
    {
        var transformer = new RecordingTransformer();
        using var importer = XlsxBuilder.Create(["Amount"], [2.5]);
        importer.Configuration.RegisterClassMap<NumberRow>(c => c.Map(r => r.Amount).WithTransformers(transformer));
        Assert.Equal(2.5, importer.ReadSheet().ReadRows<NumberRow>().Single().Amount);
        Assert.Equal([CellKind.Number], transformer.Kinds);
    }

    [Fact]
    public void ReadRows_TransformerReturnsEmpty_UsesEmptyFallback()
    {
        using var importer = Helpers.GetCsvImporter("Amount\nn/a\n");
        importer.Configuration.RegisterClassMap<NumberRow>(c => c.Map(r => r.Amount).WithTransformers(new NotApplicableTransformer()).WithEmptyFallback(-1.0));
        Assert.Equal(-1, importer.ReadSheet().ReadRows<NumberRow>().Single().Amount);
    }

    [Fact]
    public void ReadRows_Context_ReturnsExpected()
    {
        var transformer = new RecordingTransformer();
        using var importer = Helpers.GetCsvImporter("Other,Name\nx,a\ny,b\n");
        importer.Configuration.RegisterClassMap<StringRow>(c => c.Map(r => r.Name).WithTransformers(transformer));
        importer.ReadSheet().ReadRows<StringRow>().ToArray();
        Assert.Equal(["B2", "B3"], transformer.Addresses);
        Assert.All(transformer.Members, m => Assert.Equal(nameof(StringRow.Name), m));
    }

    [Fact]
    public void ReadRows_WithTransformersAfterWithTrim_ReplacesTrim()
    {
        using var importer = XlsxBuilder.Create(["Name"], ["  abc  "]);
        importer.Configuration.RegisterClassMap<StringRow>(c => c.Map(r => r.Name).WithTrim().WithTransformers(new UpperTransformer()));
        Assert.Equal("  ABC  ", importer.ReadSheet().ReadRows<StringRow>().Single().Name);
    }

    [Fact]
    public void ReadRows_TrimAttributeAndTransformerAttribute_BothApply()
    {
        using var importer = XlsxBuilder.Create(["Name"], ["  abc  "]);
        Assert.Equal("ABC", importer.ReadSheet().ReadRows<TrimmedRow>().Single().Name);
    }

    public class TrimmedRow
    {
        [ExcelTrimString]
        [ExcelTransformer(typeof(UpperTransformer))]
        public string? Name { get; set; }
    }

    [Fact]
    public void ReadRows_Elements_Transforms()
    {
        using var importer = XlsxBuilder.Create(["Values"], ["a; b"]);
        importer.Configuration.RegisterClassMap<ListRow>(c => c.Map(r => r.Values).WithSeparators(';').WithElementMap(e => e.WithTransformers(new TrimTransformer(), new UpperTransformer())));
        Assert.Equal(["A", "B"], importer.ReadSheet().ReadRows<ListRow>().Single().Values);
    }

    public class StringRow
    {
        public string? Name { get; set; }
    }

    public class NumberRow
    {
        public double Amount { get; set; }
    }

    public class ListRow
    {
        public List<string> Values { get; set; } = [];
    }

    public class AttributeRow
    {
        [ExcelTransformer(typeof(UpperTransformer))]
        public string? Name { get; set; }

        [ExcelTransformer(typeof(FrenchNumberTransformer))]
        public double Amount { get; set; }
    }

    [Fact]
    public void ReadRows_ObjectPipeline_Transforms()
    {
        // A mapping dictionary isn't a converter, so values go through the pipeline as objects.
        using var importer = XlsxBuilder.Create(["Name"], ["  abc  "]);
        importer.Configuration.RegisterClassMap<StringRow>(c => c.Map(r => r.Name).WithTransformers(new TrimTransformer()).WithMapping(new Dictionary<string, string> { ["x"] = "y" }));
        Assert.Equal("abc", importer.ReadSheet().ReadRows<StringRow>().Single().Name);
    }

    public class TrimTransformer : ICellTransformer
    {
        public CellValue Transform(CellValue value, CellContext context) => value.Kind == CellKind.Text ? new CellValue(value.Text!.Trim()) : value;
    }

    public class UpperTransformer : ICellTransformer
    {
        public CellValue Transform(CellValue value, CellContext context) => value.Kind == CellKind.Text ? new CellValue(value.Text!.ToUpperInvariant()) : value;
    }

    public class FrenchNumberTransformer : ICellTransformer
    {
        public CellValue Transform(CellValue value, CellContext context)
            => value.Kind == CellKind.Text && double.TryParse(value.Text, NumberStyles.Number, new CultureInfo("fr-FR") { NumberFormat = { NumberGroupSeparator = " " } }, out var number)
                ? new CellValue(number)
                : value;
    }

    public class NotApplicableTransformer : ICellTransformer
    {
        public CellValue Transform(CellValue value, CellContext context) => value.Text == "n/a" ? new CellValue((string?)null) : value;
    }

    private sealed class RecordingTransformer : ICellTransformer
    {
        public List<CellKind> Kinds { get; } = [];
        public List<string> Addresses { get; } = [];
        public List<string?> Members { get; } = [];

        public CellValue Transform(CellValue value, CellContext context)
        {
            Kinds.Add(value.Kind);
            Addresses.Add(context.CellAddress);
            Members.Add(context.Member?.Name);
            return value;
        }
    }
}
