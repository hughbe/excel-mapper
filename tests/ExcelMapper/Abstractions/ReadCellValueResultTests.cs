﻿using Xunit;

namespace ExcelMapper.Abstractions.Tests;

public class ReadCellResultTests
{
    [Fact]
    public void Ctor_Default()
    {
        var result = new ReadCellResult();
        Assert.Equal(0, result.ColumnIndex);
        Assert.Null(result.StringValue);
    }

    [Theory]
    [InlineData(-1, null)]
    [InlineData(0, "")]
    [InlineData(2, "abc")]
    public void Ctor_ColumnIndex_StringValue(int columnIndex, string? stringValue)
    {
        var result = new ReadCellResult(columnIndex, stringValue);
        Assert.Equal(columnIndex, result.ColumnIndex);
        Assert.Equal(stringValue, result.StringValue);
    }
}
