namespace ExcelMapper.Tests;

public class ExcelColumnNameAttributeTests
{
    [Theory]
    [InlineData("columnname")]
    [InlineData("ColumnName")]
    [InlineData("  ColumnName  ")]
    [InlineData("    ")]
    public void Ctor_String(string name)
    {
        var attribute = new ExcelColumnNameAttribute(name);
        Assert.Equal(name, attribute.Name);
        Assert.Equal(StringComparison.OrdinalIgnoreCase, attribute.Comparison);
    }

    [Fact]
    public void Ctor_NullColumnName_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnName", () => new ExcelColumnNameAttribute(null!));
    }

    [Fact]
    public void Ctor_EmptyColumnName_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnName", () => new ExcelColumnNameAttribute(string.Empty));
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
        var attribute = new ExcelColumnNameAttribute("Name")
        {
            Comparison = value
        };
        Assert.Equal(value, attribute.Comparison);

        // Set same.
        attribute.Comparison = value;
        Assert.Equal(value, attribute.Comparison);
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture - 1)]
    [InlineData(StringComparison.OrdinalIgnoreCase + 1)]
    public void Comparison_SetInvalidValue_ThrowsArgumentOutOfRangeException(StringComparison comparison)
    {
        var attribute = new ExcelColumnNameAttribute("Name");
        Assert.Throws<ArgumentOutOfRangeException>("value", () => attribute.Comparison = comparison);
    }
}
