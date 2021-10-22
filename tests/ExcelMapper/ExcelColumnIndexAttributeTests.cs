using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelColumnIndexAttributeTests
    {
        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void Ctor_Int(int index)
        {
            var attribute = new ExcelColumnIndexAttribute(index);
            Assert.Equal(index, attribute.Index);
        }

        [Fact]
        public void Ctor_InvalidIndex_ThrowsArgumentOutOfRangeException()
        {
            Assert.Throws<ArgumentOutOfRangeException>("index", () => new ExcelColumnIndexAttribute(-1));
        }
        
        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void Index_Set_GetReturnsExpectedInt(int index)
        {
            var attribute = new ExcelColumnIndexAttribute(10)
            {
                Index = index
            };
            Assert.Equal(index, attribute.Index);
        }
        
        [Fact]
        public void Index_Set_GetInvalidIndex_ThrowsArgumentOutOfRangeException()
        {
            var attribute = new ExcelColumnIndexAttribute(1);
            Assert.Throws<ArgumentOutOfRangeException>("value", () => attribute.Index = -1);
        }
    }
}
