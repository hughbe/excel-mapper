using System;
using Xunit;

namespace ExcelMapper.Mappings.Tests
{
    public class OptionalPropertyMapperTests
    {
        [Fact]
        public void Ctor_Mapper()
        {
            var innerMapper = new ColumnPropertyMapper("ColumnName");
            var mapper = new OptionalPropertyMapper(innerMapper);
            Assert.Same(innerMapper, mapper.Mapper);
        }

        [Fact]
        public void Ctor_NullMapper_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("mapper", () => new OptionalPropertyMapper(null));
        }

        [Fact]
        public void Mapper_SetValid_GetReturnsExpected()
        {
            var mapper = new OptionalPropertyMapper(new ColumnPropertyMapper("ColumnName1"));

            var innerMapper = new ColumnPropertyMapper("ColumnName2");
            mapper.Mapper = innerMapper;
            Assert.Same(innerMapper, mapper.Mapper);
        }

        [Fact]
        public void Mapper_SetNull_ThrowsArgumentNullException()
        {
            var innerMapper = new ColumnPropertyMapper("ColumnName");
            var mapper = new OptionalPropertyMapper(innerMapper);

            Assert.Throws<ArgumentNullException>("value", () => mapper.Mapper = null);
        }
    }
}
