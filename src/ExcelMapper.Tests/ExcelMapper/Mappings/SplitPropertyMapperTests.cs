using System;
using Xunit;

namespace ExcelMapper.Mappings.Tests
{
    public class SplitPropertyMapperTests
    {
        [Fact]
        public void Ctor_Mapper()
        {
            var innerMapper = new ColumnPropertyMapper("ColumnName");
            var mapper = new SplitPropertyMapper(innerMapper);
            Assert.Same(innerMapper, mapper.Mapper);

            Assert.Equal(StringSplitOptions.None, mapper.Options);
            Assert.Equal(new char[] { ',' }, mapper.Separators);
        }

        [Fact]
        public void Ctor_NullMapper_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("mapper", () => new SplitPropertyMapper(null));
        }

        [Theory]
        [InlineData(new char[] { ',' })]
        [InlineData(new char[] { ',', ';' })]
        public void Separators_SetValid_GetReturnsExpected(char[] separators)
        {
            var innerMapper = new ColumnPropertyMapper("ColumnName");
            var mapper = new SplitPropertyMapper(innerMapper) { Separators = separators };

            Assert.Same(separators, mapper.Separators);
        }

        [Fact]
        public void Separators_SetNull_ThrowsArgumentNullException()
        {
            var innerMapper = new ColumnPropertyMapper("ColumnName");
            var mapper = new SplitPropertyMapper(innerMapper);

            Assert.Throws<ArgumentNullException>("value", () => mapper.Separators = null);
        }

        [Fact]
        public void Separators_SetEmpty_ThrowsArgumentException()
        {
            var innerMapper = new ColumnPropertyMapper("ColumnName");
            var mapper = new SplitPropertyMapper(innerMapper);

            Assert.Throws<ArgumentException>("value", () => mapper.Separators = new char[0]);
        }

        [Theory]
        [InlineData(StringSplitOptions.None - 1)]
        [InlineData(StringSplitOptions.None)]
        [InlineData(StringSplitOptions.RemoveEmptyEntries)]
        [InlineData(StringSplitOptions.RemoveEmptyEntries + 1)]
        public void Options_Set_GetReturnsExpected(StringSplitOptions options)
        {
            var innerMapper = new ColumnPropertyMapper("ColumnName");
            var mapper = new SplitPropertyMapper(innerMapper) { Options = options };

            Assert.Equal(options, mapper.Options);
        }

        [Fact]
        public void Mapper_SetValid_GetReturnsExpected()
        {
            var innerMapper = new ColumnPropertyMapper("ColumnName1");
            var mapper = new SplitPropertyMapper(new ColumnPropertyMapper("ColumnName2")) { Mapper = innerMapper };

            Assert.Same(innerMapper, mapper.Mapper);
        }

        [Fact]
        public void Mapper_SetNull_ThrowsArgumentNullException()
        {
            var innerMapper = new ColumnPropertyMapper("ColumnName");
            var mapper = new SplitPropertyMapper(innerMapper);

            Assert.Throws<ArgumentNullException>("value", () => mapper.Mapper = null);
        }
    }
}
