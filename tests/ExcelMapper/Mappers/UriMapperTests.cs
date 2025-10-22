using System.Runtime.InteropServices;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers.Tests;

public class UriMapperTests
{
    public static IEnumerable<object[]> Map_TestData()
    {
        yield return new object[] { "http://microsoft.com", new Uri("http://microsoft.com") };
    }

    [Theory]
    [MemberData(nameof(Map_TestData))]
    public void Map_ValidStringValue_ReturnsSuccess(string stringValue, Uri expected)
    {
        var map = new UriMapper();
        var result = map.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
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
        var map = new UriMapper();
        var result = map.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.NotNull(result.Exception);
    }

    [Theory]
    [InlineData("/relative")]
    public void Map_InvalidStringValueWindows_ReturnsInvalid(string stringValue)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return;
        }

        Map_InvalidStringValue_ReturnsInvalid(stringValue);
    }
}
