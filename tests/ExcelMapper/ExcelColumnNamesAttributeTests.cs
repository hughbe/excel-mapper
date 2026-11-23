namespace ExcelMapper.Tests;

public class ExcelColumnNamesAttributeTests
{
    public static IEnumerable<object[]> Ctor_ParamsString_TestData()
    {
        yield return new object[] { new string[] { "ColumnName1" } };
        yield return new object[] { new string[] { "ColumnName1", "ColumnName2" } };
        yield return new object[] { new string[] { " " } };
        yield return new object[] { new string[] { "ColumnName", "ColumnName" } };
    }

    [Theory]
    [MemberData(nameof(Ctor_ParamsString_TestData))]
    public void Ctor_ParamsString(string[] columnNames)
    {
        var attribute = new ExcelColumnNamesAttribute(columnNames);
        Assert.Equal(StringComparison.OrdinalIgnoreCase, attribute.Comparison);
        Assert.Same(columnNames, attribute.Names);
    }

    [Fact]
    public void Ctor_NullColumnNames_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnNames", () => new ExcelColumnNamesAttribute(null!));
    }

    [Fact]
    public void Ctor_EmptyColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ExcelColumnNamesAttribute([]));
    }

    [Fact]
    public void Ctor_NullValueInColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ExcelColumnNamesAttribute([null!]));
    }

    [Fact]
    public void Ctor_EmptyValueInColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ExcelColumnNamesAttribute([""]));
    }

    public static IEnumerable<object[]> Names_Set_TestData()
    {
        yield return new object[] { new string[] { "ColumnName1" } };
        yield return new object[] { new string[] { "ColumnName1", "ColumnName2" } };
        yield return new object[] { new string[] { " " } };
        yield return new object[] { new string[] { "ColumnName", "ColumnName" } };
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture)]
    [InlineData(StringComparison.CurrentCultureIgnoreCase)]
    [InlineData(StringComparison.InvariantCulture)]
    [InlineData(StringComparison.InvariantCultureIgnoreCase)]
    [InlineData(StringComparison.Ordinal)]
    [InlineData(StringComparison.OrdinalIgnoreCase)]
    public void Comparison_Set_GetReturnsExpected(StringComparison value)
    {
        var attribute = new ExcelColumnNamesAttribute("ColumnName")
        {
            Comparison = value
        };
        Assert.Equal(value, attribute.Comparison);

        // Set same.
        attribute.Comparison = value;
        Assert.Equal(value, attribute.Comparison);
    }
}
