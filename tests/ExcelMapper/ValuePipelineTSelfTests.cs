using System.Linq;

namespace ExcelMapper.Tests;

public class ValuePipelineTSelfTests : ExcelClassMap<ValuePipelineTSelfTests.Row>
{
    public class Row
    {
        public string? Text { get; set; }
        public int Number { get; set; }
        public int? NullableNumber { get; set; }
        public object? Object { get; set; }
    }

    [Fact]
    public void Methods_ValueMap_ReturnTheMap()
    {
        var map = Map(r => r.Number);
        Assert.Same(map, map.WithTrim());
        Assert.Same(map, map.WithTransformers(new UpperCellTransformer()));
        Assert.Same(map, map.WithMapping(new Dictionary<string, int> { { "one", 1 } }));
        Assert.Same(map, map.WithConverter(value => 1));
        Assert.Same(map, map.WithFallback(1));
        Assert.Same(map, map.WithEmptyFallback(1));
        Assert.Same(map, map.WithInvalidFallback(1));
        Assert.Same(map, map.WithThrowingFallback());
        Assert.Same(map, map.WithThrowingEmptyFallback());
        Assert.Same(map, map.WithThrowingInvalidFallback());

        // The methods return a ValueMap<T>, so the map's own methods can follow.
        Assert.Same(map, map.WithTrim().MakeOptional());
    }

    [Fact]
    public void Methods_ElementMap_ReturnTheMap()
    {
        var map = Helpers.ElementsOf(new ExcelClassMap<CollectionRow>().Map(r => r.Numbers));
        Assert.Same(map, map.WithTrim());
        Assert.Same(map, map.WithTransformers(new UpperCellTransformer()));
        Assert.Same(map, map.WithMapping(new Dictionary<string, int> { { "one", 1 } }));
        Assert.Same(map, map.WithConverter(value => 1));
        Assert.Same(map, map.WithFallback(1));
        Assert.Same(map, map.WithEmptyFallback(1));
        Assert.Same(map, map.WithInvalidFallback(1));
        Assert.Same(map, map.WithThrowingFallback());
        Assert.Same(map, map.WithThrowingEmptyFallback());
        Assert.Same(map, map.WithThrowingInvalidFallback());
    }

    public class CollectionRow
    {
        public int[] Numbers { get; set; } = default!;
    }

    [Fact]
    public void Fallbacks_NullForNullableMember_Success()
    {
        var map = Map(r => r.NullableNumber);
        Assert.Same(map, map.WithFallback(null));
        Assert.Same(map, map.WithEmptyFallback(null));
        Assert.Same(map, map.WithInvalidFallback(null));

        var textMap = Map(r => r.Text);
        Assert.Same(textMap, textMap.WithFallback(null));
        Assert.Same(textMap, textMap.WithEmptyFallback(null));
        Assert.Same(textMap, textMap.WithInvalidFallback(null));
    }

    [Fact]
    public void WithTransformers_NullTransformers_ThrowsArgumentNullException()
    {
        var map = Map(r => r.Text);
        Assert.Throws<ArgumentNullException>("transformers", () => map.WithTransformers(null!));
        Assert.Throws<ArgumentException>("transformers", () => map.WithTransformers(new UpperCellTransformer(), null!));
    }

    [Fact]
    public void WithMapping_InvalidArguments_Throws()
    {
        var map = Map(r => r.Text);
        Assert.Throws<ArgumentNullException>("mappingDictionary", () => map.WithMapping((Dictionary<string, string>)null!));
        Assert.Throws<ArgumentOutOfRangeException>("behavior", () => map.WithMapping(new Dictionary<string, string>(), behavior: (MappingDictionaryBehavior)99));

        // Keys that are the same with the comparer.
        var duplicates = new Dictionary<string, string> { { "a", "1" }, { "A", "2" } };
        Assert.Throws<ArgumentException>(() => map.WithMapping(duplicates, StringComparer.OrdinalIgnoreCase));
    }

    [Fact]
    public void WithConverter_NullConverter_ThrowsArgumentNullException()
    {
        var map = Map(r => r.NullableNumber);
        Assert.Throws<ArgumentNullException>("converter", () => map.WithConverter((Func<string?, int?>)null!));
    }

    [Fact]
    public void ReadRows_WithMappingDerivedValueType_ReturnsExpected()
    {
        // A string? member mapped with a Dictionary<string, string>, and an object member with a
        // Dictionary<string, int>.
        using var importer = Helpers.GetCsvImporter("Text,Object\na,one\nB,two\nc,three\n");
        importer.Configuration.RegisterClassMap<Row>(c =>
        {
            c.Map(r => r.Text)
                .WithMapping(new Dictionary<string, string> { { "a", "first" }, { "b", "second" } }, StringComparer.OrdinalIgnoreCase);
            c.Map(r => r.Object)
                .WithMapping(new Dictionary<string, int> { { "one", 1 }, { "two", 2 } });
        });

        var rows = importer.ReadSheet().ReadRows<Row>().ToArray();
        Assert.Equal(["first", "second", "c"], rows.Select(r => r.Text));
        Assert.Equal([1, 2, "three"], rows.Select(r => r.Object));
    }

    private sealed class UpperCellTransformer : ICellTransformer
    {
        public CellValue Transform(CellValue value, CellContext context) => new(value.Text?.ToUpperInvariant());
    }
}
