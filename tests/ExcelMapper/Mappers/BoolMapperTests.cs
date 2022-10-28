﻿using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests
{
    public class BoolMapperTests
    {
        [Theory]
        [InlineData("1", true)]
        [InlineData("0", false)]
        [InlineData("true", true)]
        [InlineData("false", false)]
        public void GetProperty_ValidStringValue_ReturnsSuccess(string stringValue, bool expected)
        {
            var item = new BoolMapper();

            CellValueMapperResult result = item.MapCell(new ExcelCell(null, -1, -1), CellValueMapperResult.Success(stringValue), null);
            Assert.True(result.Succeeded);
            Assert.Equal(expected, result.Value);
            Assert.Null(result.Exception);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("invalid")]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(string stringValue)
        {
            var item = new BoolMapper();

            CellValueMapperResult result = item.MapCell(new ExcelCell(null, -1, -1), CellValueMapperResult.Success(stringValue), null);
            Assert.False(result.Succeeded);
            Assert.Null(result.Value);
            Assert.NotNull(result.Exception);
        }
    }
}
