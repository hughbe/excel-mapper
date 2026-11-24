using System.Linq;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Tests;

public class ValueMapTTests
{
    [Fact]
    public void MakeOptional_Invoke_SetsOptional()
    {
        var map = CreateMap();
        Assert.False(map.Optional);
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
    }

    [Fact]
    public void MakePreserveFormatting_Invoke_SetsPreserveFormatting()
    {
        var map = CreateMap();
        Assert.False(map.PreserveFormatting);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
    }

    private static ValueMap<string> CreateMap() => new ExcelClassMap<TestClass>().Map(t => t.Value);

    [Fact]
    public void ColumnMethods_InvalidArguments_Throw()
    {
        var map = CreateMap();
        Assert.Throws<ArgumentNullException>("columnName", () => map.WithColumnName(null!));
        Assert.Throws<ArgumentException>("columnName", () => map.WithColumnName(""));
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => map.WithColumnName("Name", (StringComparison)99));
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNames((string[])null!));
        Assert.Throws<ArgumentNullException>("predicate", () => map.WithColumnMatching((Func<string, bool>)null!));
        Assert.Throws<ArgumentNullException>("regex", () => map.WithColumnMatching((System.Text.RegularExpressions.Regex)null!));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => map.WithColumnIndex(-1));
        Assert.Throws<ArgumentNullException>("columnLetter", () => map.WithColumnLetter(null!));
        Assert.Throws<ArgumentException>("columnLetter", () => map.WithColumnLetter("1"));
        Assert.Throws<ArgumentNullException>("selector", () => map.WithColumns(null!));
    }

    [Fact]
    public void WithColumnNames_ArrayChangedAfterwards_ReadsGivenColumns()
    {
        var names = new[] { "A" };
        using var importer = Helpers.GetCsvImporter("A,B\na,b\n");
        importer.Configuration.RegisterClassMap<TestClass>(c => c.Map(t => t.Value).WithColumnNames(names));
        names[0] = "B";

        Assert.Equal("a", importer.ReadSheet().ReadRows<TestClass>().Single().Value);
    }

    [Fact]
    public void MakeOptional_MakePreserveFormatting_SetProperties()
    {
        var map = CreateMap();
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
    }

    private sealed class LastColumnSelector : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context) => [context.NumberOfColumns - 1];

        public string Describe(ColumnSelectorContext context) => "the last column";
    }

    private class TestClass
    {
        public string Value { get; set; } = default!;
#pragma warning disable 0649
        public string _field = default!;
#pragma warning restore 0649

        public event EventHandler Event { add { } remove { } }
    }
}
