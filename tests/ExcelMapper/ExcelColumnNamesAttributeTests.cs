using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Tests;

public class ExcelColumnNamesAttributeTests
{
    public static IEnumerable<object[]> Ctor_ParamsString()
    {
        yield return new object[] { new string[] { "ColumnName1" } };
        yield return new object[] { new string[] { "ColumnName1", "ColumnName2" } };
        yield return new object[] { new string[] { " " } };
        yield return new object[] { new string[] { "ColumnName", "ColumnName" } };
    }

    [Theory]
    [MemberData(nameof(Ctor_ParamsString))]
    public void Ctor_ColumnNames(string[] columnNames)
    {
        var attribute = new ExcelColumnNamesAttribute(columnNames);
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
}
