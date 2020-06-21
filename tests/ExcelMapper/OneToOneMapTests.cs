using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class OneToOneMapTests
    {
        [Fact]
        public void Ctor_ISingleCellValueReader()
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap(reader);
            Assert.Same(reader, map.CellReader);
            Assert.False(map.Optional);
        }

        [Fact]
        public void Ctor_NullReader_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("reader", () => new SubOneToOneMap(null));
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Optional_Set_GetReturnsExpected(bool value)
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap(reader)
            {
                Optional = value
            };
            Assert.Equal(value, map.Optional);

            // Set same.
            map.Optional = value;
            Assert.Equal(value, map.Optional);

            // Set different.
            map.Optional = !value;
            Assert.Equal(!value, map.Optional);
        }

        public static IEnumerable<object[]> CellReader_Set_TestData()
        {
            yield return new object[] { new ColumnNameValueReader("Column") };
        }

        [Theory]
        [MemberData(nameof(CellReader_Set_TestData))]
        public void CellReader_SetValid_GetReturnsExpected(ISingleCellValueReader value)
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap(reader)
            {
                CellReader = value
            };
            Assert.Same(value, map.CellReader);

            // Set same.
            map.CellReader = value;
            Assert.Same(value, map.CellReader);
        }

        [Fact]
        public void CellReader_SetNull_ThrowsArgumentNullException()
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap(reader);

            Assert.Throws<ArgumentNullException>("value", () => map.CellReader = null);
        }

        private class TestClass
        {
            public string Value { get; set; }
        }

        private class SubOneToOneMap : OneToOneMap
        {
            public SubOneToOneMap(ISingleCellValueReader reader) : base(reader)
            {
            }

            public override bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo member, out object value)
            {
                throw new NotImplementedException();
            }
        }
    }
}
