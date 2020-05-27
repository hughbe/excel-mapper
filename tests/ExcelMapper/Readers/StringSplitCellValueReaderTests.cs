using System;
using System.Collections.Generic;
using ExcelDataReader;
using Xunit;

namespace ExcelMapper.Readers.Tests
{
    public class StringSplitCellValueReaderTests
    {
        [Fact]
        public void Ctor_CellReader()
        {
            var innerReader = new ColumnNameValueReader("ColumnName");
            var reader = new StringSplitCellValueReader(innerReader);
            Assert.Same(innerReader, reader.CellReader);

            Assert.Equal(StringSplitOptions.None, reader.Options);
            Assert.Equal(new string[] { "," }, reader.Separators);
        }

        public static IEnumerable<object[]> Separators_Set_TestData()
        {
            yield return new object[] { new string[] { "," } };
            yield return new object[] { new string[] { ",", ";" } };
        }

        [Theory]
        [MemberData(nameof(Separators_Set_TestData))]
        public void Separators_SetValid_GetReturnsExpected(string[] separators)
        {
            var reader = new StringSplitCellValueReader(new ColumnNameValueReader("ColumnName")) { Separators = separators };
            Assert.Same(separators, reader.Separators);
        }

        [Fact]
        public void Separators_SetNull_ThrowsArgumentNullException()
        {
            var reader = new StringSplitCellValueReader(new ColumnNameValueReader("ColumnName"));
            Assert.Throws<ArgumentNullException>("value", () => reader.Separators = null);
        }

        [Fact]
        public void Separators_SetEmpty_ThrowsArgumentException()
        {
            var reader = new StringSplitCellValueReader(new ColumnNameValueReader("ColumnName"));
            Assert.Throws<ArgumentException>("value", () => reader.Separators = new string[0]);
        }
    }
}
