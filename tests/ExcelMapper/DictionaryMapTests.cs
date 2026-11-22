using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelMapper.Tests;

public class DictionaryMapTests
{
    private const string Csv = "Name,Q1,Q2\na,1,2\n";

    private static Dictionary<string, int> Read(Func<DictionaryMap<string, int>, DictionaryMap<string, int>> configure, string csv = Csv)
    {
        var classMap = new ExcelClassMap<Row>();
        var map = classMap.Map(r => r.Values);
        Assert.Same(map, configure(map));
        using var importer = Helpers.GetCsvImporter(csv);
        importer.Configuration.RegisterClassMap(classMap);
        return importer.ReadSheet().ReadRows<Row>().Single().Values;
    }

    private static DictionaryMap<string, int> CreateMap() => new ExcelClassMap<Row>().Map(r => r.Values);

    [Fact]
    public void Map_Default_ReturnsExpected()
    {
        var map = CreateMap();
        Assert.False(map.Optional);
        Assert.False(map.PreserveFormatting);
        Assert.NotNull(Helpers.ValuesOf(map));
    }

    [Fact]
    public void WithColumnNames_Invoke_ReadsColumns()
    {
        Assert.Equal(new Dictionary<string, int> { ["Q2"] = 2, ["Q1"] = 1 }, Read(m => m.WithColumnNames("Q2", "Q1")));
        Assert.Equal(new Dictionary<string, int> { ["Q1"] = 1 }, Read(m => m.WithColumnNames(["q1"], StringComparison.OrdinalIgnoreCase)));
        Assert.Equal(new Dictionary<string, int> { ["Q1"] = 1, ["Q2"] = 2 }, Read(m => m.WithColumnNames((IEnumerable<string>)new List<string> { "Q1", "Q2" })));
        Assert.Equal(new Dictionary<string, int> { ["Q1"] = 1 }, Read(m => m.WithColumnNames((IEnumerable<string>)new List<string> { "Q1" }, StringComparison.Ordinal)));
        Assert.Throws<ArgumentNullException>("columnNames", () => CreateMap().WithColumnNames((IReadOnlyList<string>)null!));
        Assert.Throws<ArgumentNullException>("columnNames", () => CreateMap().WithColumnNames((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithColumnsMatching_Invoke_ReadsMatchingColumns()
    {
        Assert.Equal(new Dictionary<string, int> { ["Q1"] = 1, ["Q2"] = 2 }, Read(m => m.WithColumnsMatching(n => n.StartsWith("Q"))));
        Assert.Equal(new Dictionary<string, int> { ["Q2"] = 2 }, Read(m => m.WithColumnsMatching(new Regex("2$"))));
        Assert.Throws<ArgumentNullException>("predicate", () => CreateMap().WithColumnsMatching((Func<string, bool>)null!));
        Assert.Throws<ArgumentNullException>("regex", () => CreateMap().WithColumnsMatching((Regex)null!));
    }

    [Fact]
    public void WithColumnIndices_Invoke_ReadsColumns()
    {
        Assert.Equal(new Dictionary<string, int> { ["Q2"] = 2 }, Read(m => m.WithColumnIndices(2)));
        Assert.Equal(new Dictionary<string, int> { ["Q1"] = 1, ["Q2"] = 2 }, Read(m => m.WithColumnIndices((IEnumerable<int>)new List<int> { 1, 2 })));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => CreateMap().WithColumnIndices(-1));
    }

    [Fact]
    public void WithColumns_Invoke_ReadsChosenColumns()
    {
        Assert.Equal(new Dictionary<string, int> { ["Q1"] = 1 }, Read(m => m.WithColumns(new PositionsSelector(1))));
        Assert.Throws<ArgumentNullException>("selector", () => CreateMap().WithColumns(null!));
    }

    [Fact]
    public void WithValueMap_ValidMap_Success()
    {
        var map = CreateMap();
        var newPipeline = Helpers.ValuesOf(CreateMap());
        Assert.Same(map, map.WithValueMap(v =>
        {
            Assert.Same(v, Helpers.ValuesOf(map));
            return newPipeline;
        }));
        Assert.Same(newPipeline, Helpers.ValuesOf(map));
        Assert.Throws<ArgumentNullException>("valueMap", () => map.WithValueMap(null!));
        Assert.Throws<ArgumentNullException>("valueMap", () => map.WithValueMap(_ => null!));
    }

    [Fact]
    public void MakeOptional_MissingColumns_KeepsValue()
    {
        var map = CreateMap();
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
        Assert.Equal(new Dictionary<string, int> { ["initial"] = 0 }, Read(m => m.WithColumnNames("Missing").MakeOptional()));
    }

    [Theory]
    [InlineData("Missing")]
    [InlineData("Missing1", "Missing2")]
    public void TryGetValue_MissingColumns_ThrowsExcelMappingException(params string[] columnNames)
    {
        var exception = Assert.Throws<ExcelMappingException>(() => Read(m => m.WithColumnNames(columnNames)));
        Assert.Contains("Missing", exception.Message);
        Assert.Throws<ExcelMappingException>(() => Read(m => m.WithColumnIndices(10)));
    }

    [Fact]
    public void TryGetValue_SheetWithoutHeading_ThrowsExcelMappingException()
    {
        var classMap = new ExcelClassMap<Row>();
        classMap.Map(r => r.Values).WithColumnIndices(1);
        using var importer = Helpers.GetCsvImporter(Csv);
        importer.Configuration.RegisterClassMap(classMap);
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Row>().ToArray());
    }

    public class Row
    {
        public Dictionary<string, int> Values { get; set; } = new() { ["initial"] = 0 };
    }

    private sealed class PositionsSelector(params int[] positions) : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context) => positions;

        public string Describe(ColumnSelectorContext context) => "columns";
    }
}
