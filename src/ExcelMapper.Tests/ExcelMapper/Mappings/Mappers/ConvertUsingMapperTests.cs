using System;
using Xunit;

namespace ExcelMapper.Mappings.Mappers.Tests
{
    public class ConvertUsingMapperTests
    {
        [Fact]
        public void Ctor_Converter()
        {
            ConvertUsingMappingDelegate converter = (ReadResult readResult, ref object value) => PropertyMappingResultType.Success;
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
            ConvertUsingMappingDelegate converter = (ReadResult readResult, ref object readValue) =>
            {
                Assert.Equal(-1, readResult.ColumnIndex);
                Assert.Equal("string", readResult.StringValue);

                readValue = 10;
                return PropertyMappingResultType.Success;
            };
            var item = new ConvertUsingMapper(converter);

            object value = null;
            PropertyMappingResultType result = item.GetProperty(new ReadResult(-1, "string"), ref value);
            Assert.Equal(PropertyMappingResultType.Success, result);
            Assert.Equal(10, value);
        }
    }
}
