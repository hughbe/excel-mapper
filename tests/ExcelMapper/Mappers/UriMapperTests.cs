using System.Runtime.InteropServices;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers.Tests;

public class UriMapperTests
{
    [Fact]
    public void Ctor_Default()
    {
        var map = new UriMapper();
        Assert.Equal(UriKind.Absolute, map.UriKind);
    }

    [Theory]
    [InlineData(UriKind.Absolute)]
    [InlineData(UriKind.Relative)]
    [InlineData(UriKind.RelativeOrAbsolute)]
    public void UriKind_Set_GetReturnsExpected(UriKind value)
    {
        var map = new UriMapper
        {
            UriKind = value
        };
        Assert.Equal(value, map.UriKind);

        // Set same.
        map.UriKind = value;
        Assert.Equal(value, map.UriKind);
    }

    [Theory]
    [InlineData(UriKind.RelativeOrAbsolute - 1)]
    [InlineData(UriKind.Relative + 1)]
    public void UriKind_SetInvalid_ThrowsArgumentOutOfRangeException(UriKind value)
    {
        var map = new UriMapper();
        Assert.Throws<ArgumentOutOfRangeException>("value", () => map.UriKind = value);
    }

    public static IEnumerable<object[]> Map_TestData()
    {
        yield return new object[] { UriKind.Absolute, "http://microsoft.com", new Uri("http://microsoft.com") };
        yield return new object[] { UriKind.RelativeOrAbsolute, "http://microsoft.com", new Uri("http://microsoft.com") };
        yield return new object[] { UriKind.Relative, "/path", new Uri("/path", UriKind.Relative) };
    }

    [Theory]
    [MemberData(nameof(Map_TestData))]
    public void Map_ValidStringValue_ReturnsSuccess(UriKind uriKind, string stringValue, Uri expected)
    {
        var map = new UriMapper
        {
            UriKind = uriKind
        };
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
