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
    public void Ctor_NullName_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("name", () => new ExcelColumnNameAttribute(null!));
    }

    [Fact]
    public void Ctor_EmptyName_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentException>("name", () => new ExcelColumnNameAttribute(string.Empty));
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
    public void Name_SetEmpty_ThrowsArgumentNullException()
    {
        var attribute = new ExcelColumnNameAttribute("Name");
        Assert.Throws<ArgumentException>("value", () => attribute.Name = string.Empty);
    }
}
