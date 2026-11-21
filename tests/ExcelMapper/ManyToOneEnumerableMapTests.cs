using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelMapper.Tests;

public class ManyToOneEnumerableMapTests
{
    private const string Csv = "Name,Q1,Q2,Notes\na,1,2,x|y|z\n";

    private static int[] ReadQuarters(Func<ManyToOneEnumerableMap<int>, ManyToOneEnumerableMap<int>> configure, string csv = Csv)
    {
        var classMap = new ExcelClassMap<IntsRow>();
        var map = classMap.Map(r => r.Values);
        Assert.Same(map, configure(map));
        using var importer = Helpers.GetCsvImporter(csv);
        importer.Configuration.RegisterClassMap(classMap);
        return importer.ReadSheet().ReadRows<IntsRow>().Single().Values;
    }

    private static string[] ReadStrings(Func<ManyToOneEnumerableMap<string>, ManyToOneEnumerableMap<string>> configure, string csv = Csv)
    {
        var classMap = new ExcelClassMap<StringsRow>();
        var map = classMap.Map(r => r.Values);
        Assert.Same(map, configure(map));
        using var importer = Helpers.GetCsvImporter(csv);
        importer.Configuration.RegisterClassMap(classMap);
        return importer.ReadSheet().ReadRows<StringsRow>().Single().Values;
    }

    private static ManyToOneEnumerableMap<string> CreateMap() => new ExcelClassMap<StringsRow>().Map(r => r.Values);

    [Fact]
    public void Map_Default_SplitsColumnWithMemberName()
    {
        var map = CreateMap();
        Assert.False(map.Optional);
        Assert.False(map.PreserveFormatting);
        Assert.NotNull(map.Pipeline);
        Assert.Equal(["x", "y"], ReadStrings(m => m, "Values\n\"x,y\"\n"));
    }

    [Fact]
    public void WithColumnNames_Invoke_ReadsColumnsInOrder()
    {
        Assert.Equal([2, 1], ReadQuarters(m => m.WithColumnNames("Q2", "Q1")));
        Assert.Equal([2, 1], ReadQuarters(m => m.WithColumnNames(["q2", "q1"], StringComparison.OrdinalIgnoreCase)));
        Assert.Equal([1, 2], ReadQuarters(m => m.WithColumnNames((IEnumerable<string>)new List<string> { "Q1", "Q2" })));
        Assert.Equal([1, 2], ReadQuarters(m => m.WithColumnNames((IEnumerable<string>)new List<string> { "Q1", "Q2" }, StringComparison.Ordinal)));
    }

    [Fact]
    public void WithColumnNames_InvalidArguments_Throws()
    {
        var map = CreateMap();
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNames((IReadOnlyList<string>)null!));
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)null!));
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)null!, StringComparison.Ordinal));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames(Array.Empty<string>()));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames([null!]));
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => map.WithColumnNames(["A"], (StringComparison)99));
    }

    [Fact]
    public void WithColumnsMatching_Invoke_ReadsMatchingColumns()
    {
        Assert.Equal([1, 2], ReadQuarters(m => m.WithColumnsMatching(n => n.StartsWith("Q"))));
        Assert.Equal([1, 2], ReadQuarters(m => m.WithColumnsMatching(new Regex("^Q[0-9]$"))));
        Assert.Throws<ArgumentNullException>("predicate", () => CreateMap().WithColumnsMatching((Func<string, bool>)null!));
        Assert.Throws<ArgumentNullException>("regex", () => CreateMap().WithColumnsMatching((Regex)null!));
    }

    [Fact]
    public void WithColumnIndices_Invoke_ReadsColumnsInOrder()
    {
        Assert.Equal([2, 1], ReadQuarters(m => m.WithColumnIndices(2, 1)));
        Assert.Equal([1, 2], ReadQuarters(m => m.WithColumnIndices((IEnumerable<int>)new List<int> { 1, 2 })));
        Assert.Throws<ArgumentNullException>("columnIndices", () => CreateMap().WithColumnIndices((IReadOnlyList<int>)null!));
        Assert.Throws<ArgumentNullException>("columnIndices", () => CreateMap().WithColumnIndices((IEnumerable<int>)null!));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => CreateMap().WithColumnIndices(-1));
    }

    [Fact]
    public void WithColumns_Invoke_ReadsChosenColumns()
    {
        Assert.Equal([2, 1], ReadQuarters(m => m.WithColumns(new PositionsSelector(2, 1))));
        Assert.Throws<ArgumentNullException>("selector", () => CreateMap().WithColumns(null!));
    }

    [Fact]
    public void WithColumnName_Invoke_SplitsColumn()
    {
        Assert.Equal(["x", "y", "z"], ReadStrings(m => m.WithColumnName("notes").WithSeparators('|')));
        Assert.Throws<ExcelMappingException>(() => ReadStrings(m => m.WithColumnName("notes", StringComparison.Ordinal)));
        Assert.Throws<ArgumentNullException>("columnName", () => CreateMap().WithColumnName(null!));
        Assert.Throws<ArgumentException>("columnName", () => CreateMap().WithColumnName(""));
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => CreateMap().WithColumnName("A", (StringComparison)99));
    }

    [Fact]
    public void WithColumnIndex_Invoke_SplitsColumn()
    {
        Assert.Equal(["x", "y", "z"], ReadStrings(m => m.WithColumnIndex(3).WithSeparators("|")));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => CreateMap().WithColumnIndex(-1));
    }

    [Fact]
    public void WithColumnName_AfterWithColumnNames_SplitsColumn()
    {
        Assert.Equal(["x", "y", "z"], ReadStrings(m => m.WithColumnNames("Q1", "Q2").WithColumnName("Notes").WithSeparators('|')));
        Assert.Equal(["x", "y", "z"], ReadStrings(m => m.WithColumnNames("Q1", "Q2").WithColumnIndex(3).WithSeparators('|')));
    }

    [Fact]
    public void WithColumnName_AfterWithSeparators_KeepsSeparators()
    {
        Assert.Equal(["x", "y", "z"], ReadStrings(m => m.WithSeparators('|').WithColumnName("Notes")));
        Assert.Equal(["x", "y", "z"], ReadStrings(m => m.WithSeparators('|').WithColumnIndex(3)));
    }

    [Fact]
    public void WithSeparators_Invoke_SplitsWithSeparators()
    {
        Assert.Equal(["x", "y", "z"], ReadStrings(m => m.WithColumnName("Notes").WithSeparators('|')));
        Assert.Equal(["x", "y", "z"], ReadStrings(m => m.WithColumnName("Notes").WithSeparators((IEnumerable<char>)new List<char> { '|' })));
        Assert.Equal(["x", "y", "z"], ReadStrings(m => m.WithColumnName("Notes").WithSeparators("|")));
        Assert.Equal(["x", "y", "z"], ReadStrings(m => m.WithColumnName("Notes").WithSeparators((IEnumerable<string>)new List<string> { "|" })));
    }

    [Fact]
    public void WithSeparators_InvalidSeparators_Throws()
    {
        var map = CreateMap();
        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((char[])null!));
        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((string[])null!));
        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((IEnumerable<char>)null!));
        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((IEnumerable<string>)null!));
        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators(Array.Empty<char>()));
        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators(Array.Empty<string>()));
        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators([null!]));
        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators([""]));
    }

    [Fact]
    public void WithSeparators_MultipleColumns_ThrowsExcelMappingException()
    {
        var map = CreateMap().WithColumnNames("Q1", "Q2");
        Assert.Throws<ExcelMappingException>(() => map.WithSeparators(['c']));
        Assert.Throws<ExcelMappingException>(() => map.WithSeparators(new List<char> { 'c' }));
        Assert.Throws<ExcelMappingException>(() => map.WithSeparators([","]));
        Assert.Throws<ExcelMappingException>(() => map.WithSeparators(new List<string> { "," }));
    }

    [Fact]
    public void WithElementMap_ValidMap_Success()
    {
        var map = CreateMap();
        var newPipeline = new ValuePipeline<string>();
        Assert.Same(map, map.WithElementMap(e =>
        {
            Assert.Same(e, map.Pipeline);
            return newPipeline;
        }));
        Assert.Same(newPipeline, map.Pipeline);
        Assert.Throws<ArgumentNullException>("elementMap", () => map.WithElementMap(null!));
        Assert.Throws<ArgumentNullException>("elementMap", () => map.WithElementMap(_ => null!));
    }

    [Fact]
    public void MakeOptional_MissingColumns_KeepsValue()
    {
        var map = CreateMap();
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
        Assert.Equal(["initial"], ReadStrings(m => m.WithColumnNames("Missing").MakeOptional()));
    }

    [Theory]
    [InlineData("Missing")]
    [InlineData("Missing1", "Missing2")]
    public void TryGetValue_MissingColumnNames_ThrowsExcelMappingException(params string[] columnNames)
    {
        var exception = Assert.Throws<ExcelMappingException>(() => ReadQuarters(m => m.WithColumnNames(columnNames)));
        Assert.Contains("Missing", exception.Message);
        Assert.Equal(nameof(IntsRow.Values), exception.Member!.Name);
    }

    [Theory]
    [InlineData(10)]
    [InlineData(10, 11)]
    public void TryGetValue_MissingColumnIndices_ThrowsExcelMappingException(params int[] columnIndices)
    {
        Assert.Throws<ExcelMappingException>(() => ReadQuarters(m => m.WithColumnIndices(columnIndices)));
    }

    [Fact]
    public void TryGetValue_NoMatchingColumns_ThrowsExcelMappingException()
    {
        Assert.Throws<ExcelMappingException>(() => ReadQuarters(m => m.WithColumnsMatching(n => n == "Missing")));
    }

    [Fact]
    public void TryGetValue_NullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetCsvImporter(Csv);
        Assert.Throws<ArgumentNullException>("sheet", () => CreateMap().TryGetValue(null!, 0, importer.Reader, null, out _));
    }

    public class IntsRow
    {
        public int[] Values { get; set; } = [];
    }

    public class StringsRow
    {
        public string[] Values { get; set; } = ["initial"];
    }

    private sealed class PositionsSelector(params int[] positions) : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ExcelSheet sheet) => positions;

        public string Describe(ExcelSheet sheet) => "columns";
    }
}
