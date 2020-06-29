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
    public class OneToOneMapTTests
    {
        [Fact]
        public void Ctor_ISingleCellValueReader()
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap<int>(reader);
            Assert.Same(reader, map.CellReader);
            Assert.Empty(map.CellValueMappers);
            Assert.Same(map.CellValueMappers, map.CellValueMappers);
            Assert.Same(map.CellValueMappers, map.Pipeline.CellValueMappers);
            Assert.Empty(map.CellValueTransformers);
            Assert.Same(map.CellValueTransformers, map.CellValueTransformers);
            Assert.Same(map.CellValueTransformers, map.Pipeline.CellValueTransformers);
            Assert.False(map.Optional);
        }

        [Fact]
        public void Ctor_NullReader_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("reader", () => new SubOneToOneMap<int>(null));
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
            var map = new SubOneToOneMap<int>(reader)
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
            var map = new SubOneToOneMap<int>(reader);

            Assert.Throws<ArgumentNullException>("value", () => map.CellReader = null);
        }

        [Fact]
        public void EmptyFallback_Set_GetReturnsExpected()
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap<int>(reader);

            var fallback = new FixedValueFallback(10);
            map.EmptyFallback = fallback;
            Assert.Same(fallback, map.EmptyFallback);

            map.EmptyFallback = null;
            Assert.Null(map.EmptyFallback);
        }

        [Fact]
        public void InvalidFallback_Set_GetReturnsExpected()
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap<int>(reader);

            var fallback = new FixedValueFallback(10);
            map.InvalidFallback = fallback;
            Assert.Same(fallback, map.InvalidFallback);

            map.InvalidFallback = null;
            Assert.Null(map.InvalidFallback);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Optional_Set_GetReturnsExpected(bool value)
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap<int>(reader)
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

        [Fact]
        public void AddCellValueMapper_ValidItem_Success()
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap<int>(reader);
            var item1 = new BoolMapper();
            var item2 = new BoolMapper();

            map.AddCellValueMapper(item1);
            map.AddCellValueMapper(item2);
            Assert.Equal(new ICellValueMapper[] { item1, item2 }, map.CellValueMappers);
        }

        [Fact]
        public void AddCellValueMapper_NullItem_ThrowsArgumentNullException()
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap<int>(reader);

            Assert.Throws<ArgumentNullException>("mapper", () => map.AddCellValueMapper(null));
        }

        [Fact]
        public void RemoveCellValueMapper_Index_Success()
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap<int>(reader);
            map.AddCellValueMapper(new BoolMapper());

            map.RemoveCellValueMapper(0);
            Assert.Empty(map.CellValueMappers);
        }

        [Fact]
        public void AddCellValueTransformer_ValidTransformer_Success()
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap<int>(reader);
            var transformer1 = new TrimCellValueTransformer();
            var transformer2 = new TrimCellValueTransformer();

            map.AddCellValueTransformer(transformer1);
            map.AddCellValueTransformer(transformer2);
            Assert.Equal(new ICellValueTransformer[] { transformer1, transformer2 }, map.CellValueTransformers);
        }

        [Fact]
        public void AddCellValueTransformer_NullTransformer_ThrowsArgumentNullException()
        {
            var reader = new ColumnNameValueReader("Column");
            var map = new SubOneToOneMap<int>(reader);
            Assert.Throws<ArgumentNullException>("transformer", () => map.AddCellValueTransformer(null));
        }

        private class TestClass
        {
            public string Value { get; set; }
        }

        private class SubOneToOneMap<T> : OneToOneMap<T>
        {
            public SubOneToOneMap(ISingleCellValueReader reader) : base(reader)
            {
            }
        }
    }
}
