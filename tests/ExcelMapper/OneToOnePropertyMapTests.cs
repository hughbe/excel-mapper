using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class OneToOnePropertyMapTests
    {
        [Fact]
        public void Ctor_Member_ISingleCellValueReader()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader);
            Assert.Same(reader, propertyMap.CellReader);
            Assert.Empty(propertyMap.CellValueMappers);
            Assert.Same(propertyMap.CellValueMappers, propertyMap.CellValueMappers);
            Assert.Same(propertyMap.CellValueMappers, propertyMap.Pipeline.CellValueMappers);
            Assert.Empty(propertyMap.CellValueTransformers);
            Assert.Same(propertyMap.CellValueTransformers, propertyMap.CellValueTransformers);
            Assert.Same(propertyMap.CellValueTransformers, propertyMap.Pipeline.CellValueTransformers);
            Assert.Same(propertyInfo, propertyMap.Member);
            Assert.False(propertyMap.Optional);
        }

        [Fact]
        public void Ctor_NullMember_ThrowsArgumentNullException()
        {
            var reader = new ColumnNameValueReader("Column");
            Assert.Throws<ArgumentNullException>("member", () => new OneToOnePropertyMap(null, reader));
        }

        [Fact]
        public void Ctor_NullReader_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            Assert.Throws<ArgumentNullException>("reader", () => new OneToOnePropertyMap(propertyInfo, null));
        }

        [Fact]
        public void EmptyFallback_Set_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader);

            var fallback = new FixedValueFallback(10);
            propertyMap.EmptyFallback = fallback;
            Assert.Same(fallback, propertyMap.EmptyFallback);

            propertyMap.EmptyFallback = null;
            Assert.Null(propertyMap.EmptyFallback);
        }

        [Fact]
        public void InvalidFallback_Set_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader);

            var fallback = new FixedValueFallback(10);
            propertyMap.InvalidFallback = fallback;
            Assert.Same(fallback, propertyMap.InvalidFallback);

            propertyMap.InvalidFallback = null;
            Assert.Null(propertyMap.InvalidFallback);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Optional_Set_GetReturnsExpected(bool value)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader)
            {
                Optional = value
            };
            Assert.Equal(value, propertyMap.Optional);

            // Set same.
            propertyMap.Optional = value;
            Assert.Equal(value, propertyMap.Optional);

            // Set different.
            propertyMap.Optional = !value;
            Assert.Equal(!value, propertyMap.Optional);
        }

        [Fact]
        public void AddCellValueMapper_ValidItem_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader);
            var item1 = new BoolMapper();
            var item2 = new BoolMapper();

            propertyMap.AddCellValueMapper(item1);
            propertyMap.AddCellValueMapper(item2);
            Assert.Equal(new ICellValueMapper[] { item1, item2 }, propertyMap.CellValueMappers);
        }

        [Fact]
        public void AddCellValueMapper_NullItem_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader);

            Assert.Throws<ArgumentNullException>("mapper", () => propertyMap.AddCellValueMapper(null));
        }

        [Fact]
        public void RemoveCellValueMapper_Index_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader);
            propertyMap.AddCellValueMapper(new BoolMapper());

            propertyMap.RemoveCellValueMapper(0);
            Assert.Empty(propertyMap.CellValueMappers);
        }

        [Fact]
        public void AddCellValueTransformer_ValidTransformer_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader);
            var transformer1 = new TrimCellValueTransformer();
            var transformer2 = new TrimCellValueTransformer();

            propertyMap.AddCellValueTransformer(transformer1);
            propertyMap.AddCellValueTransformer(transformer2);
            Assert.Equal(new ICellValueTransformer[] { transformer1, transformer2 }, propertyMap.CellValueTransformers);
        }

        [Fact]
        public void AddCellValueTransformer_NullTransformer_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader);
            Assert.Throws<ArgumentNullException>("transformer", () => propertyMap.AddCellValueTransformer(null));
        }

        public static IEnumerable<object[]> CellReader_Set_TestData()
        {
            yield return new object[] { new ColumnNameValueReader("Column") };
        }

        [Theory]
        [MemberData(nameof(CellReader_Set_TestData))]
        public void CellReader_SetValid_GetReturnsExpected(ISingleCellValueReader value)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader)
            {
                CellReader = value
            };
            Assert.Same(value, propertyMap.CellReader);

            // Set same.
            propertyMap.CellReader = value;
            Assert.Same(value, propertyMap.CellReader);
        }

        [Fact]
        public void CellReader_SetNull_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var reader = new ColumnNameValueReader("Column");
            var propertyMap = new OneToOnePropertyMap(propertyInfo, reader);

            Assert.Throws<ArgumentNullException>("value", () => propertyMap.CellReader = null);
        }

        private class TestClass
        {
            public string Value { get; set; }
        }
    }
}
