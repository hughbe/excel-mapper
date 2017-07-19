using System;
using Xunit;

namespace ExcelMapper.Mappings.Readers.Tests
{
    public class OptionalCellValueReaderTests
    {
        [Fact]
        public void Ctor_InnerReader()
        {
            var innerReader = new ColumnNameValueReader("ColumnName");
            var reader = new OptionalCellValueReader(innerReader);
            Assert.Same(innerReader, reader.InnerReader);
        }

        [Fact]
        public void Ctor_NullInnerReader_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("innerReader", () => new OptionalCellValueReader(null));
        }

        [Fact]
        public void InnerReader_SetValid_GetReturnsExpected()
        {
            var reader = new OptionalCellValueReader(new ColumnNameValueReader("ColumnName1"));

            var innerReader = new ColumnNameValueReader("ColumnName2");
            reader.InnerReader = innerReader;
            Assert.Same(innerReader, reader.InnerReader);
        }

        [Fact]
        public void InnerReader_SetNull_ThrowsArgumentNullException()
        {
            var reader = new OptionalCellValueReader(new ColumnNameValueReader("ColumnName"));

            Assert.Throws<ArgumentNullException>("value", () => reader.InnerReader = null);
        }
    }
}
