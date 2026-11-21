using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers.Tests;

public class BoolMapperTests
{
    [Theory]
    [InlineData("1", true)]
    [InlineData("0", false)]
    [InlineData("true", true)]
    [InlineData("false", false)]
    public void Map_ValidStringValue_ReturnsSuccess(string stringValue, bool expected)
    {
        var item = new BoolMapper();

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
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
        var item = new BoolMapper();

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.NotNull(result.Exception);
    }

    [Theory]
    [InlineData("yes", true)]
    [InlineData("Yes", true)]
    [InlineData("YES", true)]
    [InlineData("no", false)]
    [InlineData("No", false)]
    [InlineData("NO", false)]
    public void Map_YesNoStringValue_ReturnsSuccess(string stringValue, bool expected)
    {
        var item = new BoolMapper();

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, result.Value);
        Assert.Null(result.Exception);
    }

    [Theory]
    [InlineData("y")]
    [InlineData("n")]
    [InlineData("yes ")]
    [InlineData("maybe")]
    public void Map_InvalidYesNoStringValue_ReturnsInvalid(string stringValue)
    {
        var item = new BoolMapper();

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.IsType<FormatException>(result.Exception);
    }
}
