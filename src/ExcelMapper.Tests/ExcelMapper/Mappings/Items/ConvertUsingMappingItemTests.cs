using System;
using Xunit;

namespace ExcelMapper.Mappings.Items.Tests
{
    public class ConvertUsingMappingItemTests
    {
        [Fact]
        public void Ctor_Converter()
        {
            ConvertUsingMappingDelegate converter = result => PropertyMappingResult.Success(null);
            var item = new ConvertUsingMappingItem(converter);
            Assert.Same(converter, item.Converter);
        }

        [Fact]
        public void Ctor_NullConverter_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("converter", () => new ConvertUsingMappingItem(null));
        }

        [Fact]
        public void GetProperty_ValidStringValue_ReturnsSuccess()
        {
            ConvertUsingMappingDelegate converter = mapResult =>
            {
                Assert.Equal(-1, mapResult.ColumnIndex);
                Assert.Equal("string", mapResult.StringValue);

                return PropertyMappingResult.Success(10);
            };
            var item = new ConvertUsingMappingItem(converter);

            PropertyMappingResult result = item.GetProperty(null, 0, null, new MapResult(-1, "string"));
            Assert.Equal(PropertyMappingResultType.Success, result.Type);
            Assert.Equal(10, result.Value);
        }
    }
}
