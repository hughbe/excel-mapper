using System;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests
{
    public class ConvertUsingMapperTests
    {
        [Fact]
        public void Ctor_Converter()
        {
            ConvertUsingMapperDelegate converter = (ReadCellValueResult readResult) => CellValueMapperResult.Success(1);
            var item = new ConvertUsingMapper(converter);
            Assert.Same(converter, item.Converter);
        }

        [Fact]
        public void Ctor_NullConverter_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("converter", () => new ConvertUsingMapper(null));
        }

        [Fact]
        public void GetProperty_ValidStringValue_ReturnsSuccess()
        {
            ConvertUsingMapperDelegate converter = (ReadCellValueResult readResult) =>
            {
                Assert.Equal(-1, readResult.ColumnIndex);
                Assert.Equal("string", readResult.StringValue);
                return CellValueMapperResult.Success(10);
            };
            var item = new ConvertUsingMapper(converter);
            
            CellValueMapperResult result = item.MapCellValue(new ReadCellValueResult(-1, "string"));
            Assert.True(result.Succeeded);
            Assert.Equal(10, result.Value);
            Assert.Null(result.Exception);
        }

        [Fact]
        public void GetProperty_InvalidStringValue_ReturnsSuccess()
        {
            ConvertUsingMapperDelegate converter = (ReadCellValueResult readResult) =>
            {
                Assert.Equal(-1, readResult.ColumnIndex);
                Assert.Equal("string", readResult.StringValue);
                throw new DivideByZeroException();
            };
            var item = new ConvertUsingMapper(converter);
            
            CellValueMapperResult result = item.MapCellValue(new ReadCellValueResult(-1, "string"));
            Assert.False(result.Succeeded);
            Assert.Null(result.Value);
            Assert.IsType<DivideByZeroException>(result.Exception);
        }
    }
}
