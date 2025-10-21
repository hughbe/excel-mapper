using System.Globalization;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests;

public class ParsableMapperTests
{
    [Fact]
    public void Ctor_Default()
    {
        var mapper = new ParsableMapper<nint>();
        Assert.Null(mapper.Provider);
    }

    [Fact]
    public void Provider_Set_GetReturnsExpected()
    {
        var provider = CultureInfo.CurrentCulture;
        var mapper = new ParsableMapper<nint>
        {
            Provider = provider
        };
        Assert.Same(provider, mapper.Provider);

        // Set same.
        mapper.Provider = provider;
        Assert.Same(provider, mapper.Provider);

        // Set null.
        mapper.Provider = null;
        Assert.Null(mapper.Provider);
    }

    [Theory]
    [InlineData("123", 123)]
    public void MapCellValue_ParsableValue_ReturnsExpected(string stringValue, nint expected)
    {
        var mapper = new ParsableMapper<nint>();
        var result = mapper.MapCellValue(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, Assert.IsType<nint>(result.Value));
        Assert.Null(result.Exception);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("invalid")]
    [InlineData("12/07/2017 07:57:61")]
    public void MapCellValue_InvalidStringValue_ReturnsInvalid(string? stringValue)
    {
        var item = new ParsableMapper<nint>();
        var result = item.MapCellValue(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.NotNull(result.Exception);
    }
}
