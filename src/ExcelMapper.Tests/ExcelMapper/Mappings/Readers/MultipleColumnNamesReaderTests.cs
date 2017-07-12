using System;
using Xunit;

namespace ExcelMapper.Mappings.Readers.Tests
{
    public class MultipleColumnNamesReaderTests
    {
        [Fact]
        public void Ctor_ColumnNames()
        {
            var columnNames = new string[] { "ColumnName1", "ColumnName2" };
            var reader = new MultipleColumnNamesReader(columnNames);
            Assert.Same(columnNames, reader.ColumnNames);
        }

        [Fact]
        public void Ctor_NullColumnNames_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("columnNames", () => new MultipleColumnNamesReader(null));
        }

        [Fact]
        public void Ctor_EmptyColumnNames_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnNames", () => new MultipleColumnNamesReader(new string[0]));
        }

        [Fact]
        public void Ctor_NullValueInColumnNames_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnNames", () => new MultipleColumnNamesReader(new string[] { null }));
        }
    }
}
