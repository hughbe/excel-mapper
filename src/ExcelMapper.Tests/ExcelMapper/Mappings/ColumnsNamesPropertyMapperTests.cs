using System;
using Xunit;

namespace ExcelMapper.Mappings.Tests
{
    public class ColumnsNamesPropertyMapperTests
    {
        [Fact]
        public void Ctor_ColumnNames()
        {
            var columnNames = new string[] { "ColumnName1", "ColumnName2" };
            var mapper = new ColumnsNamesPropertyMapper(columnNames);
            Assert.Equal(columnNames, mapper.ColumnNames);
        }

        [Fact]
        public void Ctor_NullColumnNames_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("columnNames", () => new ColumnsNamesPropertyMapper(null));
        }

        [Fact]
        public void Ctor_EmptyColumnNames_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnNames", () => new ColumnsNamesPropertyMapper(new string[0]));
        }

        [Fact]
        public void Ctor_NullValueInColumnNames_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnNames", () => new ColumnsNamesPropertyMapper(new string[] { null }));
        }
    }
}
