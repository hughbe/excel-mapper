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

    [Theory]
    [InlineData("columnname", StringComparison.CurrentCulture)]
    [InlineData("ColumnName", StringComparison.CurrentCultureIgnoreCase)]
    [InlineData("  ColumnName  ", StringComparison.InvariantCulture)]
    [InlineData("    ", StringComparison.InvariantCultureIgnoreCase)]
    [InlineData("ColumnName", StringComparison.Ordinal)]
    [InlineData("columnname", StringComparison.OrdinalIgnoreCase)]
    public void Ctor_String_StringComparison(string name, StringComparison comparison)
    {
        var attribute = new ExcelColumnNameAttribute(name, comparison);
        Assert.Equal(name, attribute.Name);
        Assert.Equal(comparison, attribute.Comparison);
    }

    [Fact]
    public void Ctor_NullColumnName_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnName", () => new ExcelColumnNameAttribute(null!));
        Assert.Throws<ArgumentNullException>("columnName", () => new ExcelColumnNameAttribute(null!, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Ctor_EmptyColumnName_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnName", () => new ExcelColumnNameAttribute(string.Empty));
        Assert.Throws<ArgumentException>("columnName", () => new ExcelColumnNameAttribute(string.Empty, StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture - 1)]
    [InlineData(StringComparison.OrdinalIgnoreCase + 1)]
    public void Ctor_InvalidStringComparison_ThrowsArgumentOutOfRangeException(StringComparison comparison)
    {
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => new ExcelColumnNameAttribute("Name", comparison));
    }

    [Theory]
    [InlineData("columnname")]
    [InlineData("ColumnName")]
    public void Name_Set_GetReturnsExpected(string value)
    {
        var attribute = new ExcelColumnNameAttribute("Name")
        {
            Name = value
        };
        Assert.Equal(value, attribute.Name);

        // Set same.
        attribute.Name = value;
        Assert.Equal(value, attribute.Name);
    }

    [Fact]
    public void Name_SetNull_ThrowsArgumentNullException()
    {
        var attribute = new ExcelColumnNameAttribute("Name");
        Assert.Throws<ArgumentNullException>("value", () => attribute.Name = null!);
    }

    [Fact]
    public void Name_SetEmpty_ThrowsArgumentException()
    {
        var attribute = new ExcelColumnNameAttribute("Name");
        Assert.Throws<ArgumentException>("value", () => attribute.Name = string.Empty);
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
