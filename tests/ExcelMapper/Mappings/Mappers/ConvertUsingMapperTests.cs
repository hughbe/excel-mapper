using System;
using Xunit;

namespace ExcelMapper.Mappings.Mappers.Tests
{
    public class ConvertUsingMapperTests
    {
        [Fact]
        public void Ctor_Converter()
        {
            ConvertUsingMapperDelegate converter = (ReadCellValueResult readResult, ref object value) => PropertyMapperResultType.Success;
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
            ConvertUsingMapperDelegate converter = (ReadCellValueResult readResult, ref object readValue) =>
            {
                Assert.Equal(-1, readResult.ColumnIndex);
                Assert.Equal("string", readResult.StringValue);

                readValue = 10;
                return PropertyMapperResultType.Success;
            };
            var item = new ConvertUsingMapper(converter);

            object value = null;
            PropertyMapperResultType result = item.MapCellValue(new ReadCellValueResult(-1, "string"), ref value);
            Assert.Equal(PropertyMapperResultType.Success, result);
            Assert.Equal(10, value);
        }
    }
}
