using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests;

public class VersionMapperTests
{
    public static IEnumerable<object[]> Map_TestData()
    {
        yield return new object[] { "1.0", new Version("1.0") };
    }

    [Theory]
    [MemberData(nameof(Map_TestData))]
    public void Map_ValidStringValue_ReturnsSuccess(string stringValue, Version expected)
    {
        var mapper = new VersionMapper();
        
        var result = mapper.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, result.Value);
        Assert.Null(result.Exception);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("invalid")]
    public void Map_InvalidStringValue_ReturnsInvalid(string? stringValue)
    {
        var mapper = new VersionMapper();
        var result = mapper.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.NotNull(result.Exception);
    }
}
