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

    public static IEnumerable<object[]> Ctor_IReadOnlyListString_StringComparison_TestData()
    {
        yield return new object[] { new string[] { "ColumnName1" }, StringComparison.CurrentCulture };
        yield return new object[] { new string[] { "ColumnName1", "ColumnName2" }, StringComparison.CurrentCultureIgnoreCase };
        yield return new object[] { new string[] { " " }, StringComparison.InvariantCulture };
        yield return new object[] { new string[] { "ColumnName", "ColumnName" }, StringComparison.InvariantCultureIgnoreCase };
        yield return new object[] { new string[] { "ColumnName" }, StringComparison.Ordinal };
        yield return new object[] { new string[] { "ColumnName" }, StringComparison.OrdinalIgnoreCase };
    }

    [Theory]
    [MemberData(nameof(Ctor_IReadOnlyListString_StringComparison_TestData))]
    public void Ctor_IReadOnlyListString_StringComparison(string[] columnNames, StringComparison comparison)
    {
        var attribute = new ExcelColumnNamesAttribute(columnNames, comparison);
        Assert.Same(columnNames, attribute.Names);
        Assert.Equal(comparison, attribute.Comparison);
    }

    [Fact]
    public void Ctor_NullColumnNames_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnNames", () => new ExcelColumnNamesAttribute(null!));
        Assert.Throws<ArgumentNullException>("columnNames", () => new ExcelColumnNamesAttribute(null!, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Ctor_EmptyColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ExcelColumnNamesAttribute([]));
        Assert.Throws<ArgumentException>("columnNames", () => new ExcelColumnNamesAttribute([], StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Ctor_NullValueInColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ExcelColumnNamesAttribute([null!]));
        Assert.Throws<ArgumentException>("columnNames", () => new ExcelColumnNamesAttribute([null!], StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Ctor_EmptyValueInColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ExcelColumnNamesAttribute([""]));
        Assert.Throws<ArgumentException>("columnNames", () => new ExcelColumnNamesAttribute([""], StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture - 1)]
    [InlineData(StringComparison.OrdinalIgnoreCase + 1)]
    public void Ctor_InvalidComparison_ThrowsArgumentOutOfRangeException(StringComparison comparison)
    {
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => new ExcelColumnNamesAttribute(["ColumnName"], comparison));
    }

    public static IEnumerable<object[]> Names_Set_TestData()
    {
        yield return new object[] { new string[] { "ColumnName1" } };
        yield return new object[] { new string[] { "ColumnName1", "ColumnName2" } };
        yield return new object[] { new string[] { " " } };
        yield return new object[] { new string[] { "ColumnName", "ColumnName" } };
    }

    [Theory]
    [MemberData(nameof(Names_Set_TestData))]
    public void Names_Set_GetReturnsExpected(string[] value)
    {
        var attribute = new ExcelColumnNamesAttribute("ColumnName")
        {
            Names = value
        };
        Assert.Same(value, attribute.Names);
        
        // Set same.
        attribute.Names = value;
        Assert.Same(value, attribute.Names);
    }

    [Fact]
    public void Names_SetNullValue_ThrowsArgumentNullException()
    {
        var attribute = new ExcelColumnNamesAttribute(["ColumnName"]);
        Assert.Throws<ArgumentNullException>("value", () => attribute.Names = null!);
    }

    [Fact]
    public void Names_SetEmptyValue_ThrowsArgumentException()
    {
        var attribute = new ExcelColumnNamesAttribute(["ColumnName"]);
        Assert.Throws<ArgumentException>("value", () => attribute.Names = []);
    }

    [Fact]
    public void Names_SetNullValueInValue_ThrowsArgumentException()
    {
        var attribute = new ExcelColumnNamesAttribute(["ColumnName"]);
        Assert.Throws<ArgumentException>("value", () => attribute.Names = [null!]);
    }

    [Fact]
    public void Names_SetEmptyValueInValue_ThrowsArgumentException()
    {
        var attribute = new ExcelColumnNamesAttribute(["ColumnName"]);
        Assert.Throws<ArgumentException>("value", () => attribute.Names = [""]);
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
