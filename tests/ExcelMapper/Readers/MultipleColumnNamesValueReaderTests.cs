using System;
using Xunit;

namespace ExcelMapper.Readers.Tests
{
    public class MultipleColumnNamesValueReaderTests
    {
        [Fact]
        public void Ctor_ColumnNames()
        {
            var columnNames = new string[] { "ColumnName1", "ColumnName2" };
            var reader = new MultipleColumnNamesValueReader(columnNames);
            Assert.Same(columnNames, reader.ColumnNames);
        }

        [Fact]
        public void Ctor_NullColumnNames_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("columnNames", () => new MultipleColumnNamesValueReader(null));
        }

        [Fact]
        public void Ctor_EmptyColumnNames_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnNames", () => new MultipleColumnNamesValueReader(new string[0]));
        }

        [Fact]
        public void Ctor_NullValueInColumnNames_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnNames", () => new MultipleColumnNamesValueReader(new string[] { null }));
        }
    }
}
