using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests;

public class StringMapperTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("abc")]
    public void Map_Invoke_ReturnsBegan(string? stringValue)
    {
        var item = new StringMapper();

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(stringValue, result.Value);
        Assert.Null(result.Exception);
    }
}
