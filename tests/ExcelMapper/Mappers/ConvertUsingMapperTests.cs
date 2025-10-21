using System;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests;

public class ConvertUsingMapperTests
{
    [Fact]
    public void Ctor_Converter()
    {
        ConvertUsingMapperDelegate converter = (ReadCellResult readResult) => CellMapperResult.Success(1);
        var item = new ConvertUsingMapper(converter);
        Assert.Same(converter, item.Converter);
    }

    [Fact]
    public void Ctor_NullConverter_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("converter", () => new ConvertUsingMapper(null!));
    }

    [Fact]
    public void Map_ValidStringValue_ReturnsSuccess()
    {
        ConvertUsingMapperDelegate converter = (ReadCellResult readResult) =>
        {
            Assert.Equal(0, readResult.ColumnIndex);
            Assert.Equal("string", readResult.StringValue);
            return CellMapperResult.Success(10);
        };
        var item = new ConvertUsingMapper(converter);
        
        var result = item.Map(new ReadCellResult(0, "string", preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(10, result.Value);
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_InvalidStringValue_ReturnsSuccess()
    {
        ConvertUsingMapperDelegate converter = (ReadCellResult readResult) =>
        {
            Assert.Equal(0, readResult.ColumnIndex);
            Assert.Equal("string", readResult.StringValue);
            throw new DivideByZeroException();
        };
        var item = new ConvertUsingMapper(converter);
        
        var result = item.Map(new ReadCellResult(0, "string", preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.IsType<DivideByZeroException>(result.Exception);
    }
}
