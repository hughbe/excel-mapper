using System.Linq;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Tests;

public class ValueMapTTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Optional_Set_GetReturnsExpected(bool value)
    {
        var map = CreateMap();
        map.Optional = value;
        Assert.Equal(value, map.Optional);

        // Set same.
        map.Optional = value;
        Assert.Equal(value, map.Optional);

        // Set different.
        map.Optional = !value;
        Assert.Equal(!value, map.Optional);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PreserveFormatting_Set_GetReturnsExpected(bool value)
    {
        var map = CreateMap();
        map.PreserveFormatting = value;
        Assert.Equal(value, map.PreserveFormatting);

        // Set same.
        map.PreserveFormatting = value;
        Assert.Equal(value, map.PreserveFormatting);

        // Set different.
        map.PreserveFormatting = !value;
        Assert.Equal(!value, map.PreserveFormatting);
    }

    private static ValueMap<string> CreateMap() => new ExcelClassMap<TestClass>().Map(t => t.Value);

    [Fact]
    public void ColumnMethods_InvalidArguments_Throw()
    {
        var map = CreateMap();
        Assert.Throws<ArgumentNullException>("columnName", () => map.WithColumnName(null!));
        Assert.Throws<ArgumentException>("columnName", () => map.WithColumnName(""));
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => map.WithColumnName("Name", (StringComparison)99));
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNameMatching((string[])null!));
        Assert.Throws<ArgumentNullException>("predicate", () => map.WithColumnNameMatching((Func<string, bool>)null!));
        Assert.Throws<ArgumentNullException>("regex", () => map.WithColumnNameMatching((System.Text.RegularExpressions.Regex)null!));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => map.WithColumnIndex(-1));
        Assert.Throws<ArgumentNullException>("columnLetter", () => map.WithColumnLetter(null!));
        Assert.Throws<ArgumentException>("columnLetter", () => map.WithColumnLetter("1"));
        Assert.Throws<ArgumentNullException>("selector", () => map.WithColumns(null!));
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
        public IReadOnlyList<int> SelectColumns(ExcelSheet sheet) => [sheet.NumberOfColumns - 1];

        public string Describe(ExcelSheet sheet) => "the last column";
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
