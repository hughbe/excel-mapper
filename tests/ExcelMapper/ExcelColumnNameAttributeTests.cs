using System;
using Xunit;

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
}
