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
        public void Ctor_Member_Type_EmptyValueStrategy()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));

            var propertyMap = new OneToOnePropertyMap(propertyInfo);
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
        public void EmptyFallback_Set_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo);

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
            var propertyMap = new OneToOnePropertyMap(propertyInfo);

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
            var propertyMap = new OneToOnePropertyMap(propertyInfo)
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
            var propertyMap = new OneToOnePropertyMap(propertyInfo);
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
            var propertyMap = new OneToOnePropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("mapper", () => propertyMap.AddCellValueMapper(null));
        }

        [Fact]
        public void RemoveCellValueMapper_Index_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo);
            propertyMap.AddCellValueMapper(new BoolMapper());

            propertyMap.RemoveCellValueMapper(0);
            Assert.Empty(propertyMap.CellValueMappers);
        }

        [Fact]
        public void AddCellValueTransformer_ValidTransformer_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo);
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
            var propertyMap = new OneToOnePropertyMap(propertyInfo);

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
            var propertyMap = new OneToOnePropertyMap(propertyInfo)
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
            var propertyMap = new OneToOnePropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("value", () => propertyMap.CellReader = null);
        }

        private class TestClass
        {
            public string Value { get; set; }
        }
    }
}
