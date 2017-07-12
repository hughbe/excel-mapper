using System;
using Xunit;

namespace ExcelMapper.Mappings.Readers.Tests
{
    public class OptionalColumnReaderTests
    {
        [Fact]
        public void Ctor_InnerReader()
        {
            var innerReader = new ColumnNameReader("ColumnName");
            var reader = new OptionalColumnReader(innerReader);
            Assert.Same(innerReader, reader.InnerReader);
        }

        [Fact]
        public void Ctor_NullInnerReader_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("innerReader", () => new OptionalColumnReader(null));
        }

        [Fact]
        public void InnerReader_SetValid_GetReturnsExpected()
        {
            var reader = new OptionalColumnReader(new ColumnNameReader("ColumnName1"));

            var innerReader = new ColumnNameReader("ColumnName2");
            reader.InnerReader = innerReader;
            Assert.Same(innerReader, reader.InnerReader);
        }

        [Fact]
        public void InnerReader_SetNull_ThrowsArgumentNullException()
        {
            var reader = new OptionalColumnReader(new ColumnNameReader("ColumnName"));

            Assert.Throws<ArgumentNullException>("value", () => reader.InnerReader = null);
        }
    }
}
