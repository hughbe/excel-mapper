using System;
using System.Reflection;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Mappers;
using ExcelMapper.Mappings.Readers;
using ExcelMapper.Mappings.Transformers;
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
            Assert.Same(propertyInfo, propertyMap.Member);

            Assert.Empty(propertyMap.Pipeline.CellValueMappers);
            Assert.Empty(propertyMap.Pipeline.CellValueTransformers);
        }

        [Fact]
        public void EmptyFallback_Set_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo);

            var fallback = new FixedValueFallback(10);
            propertyMap.Pipeline.EmptyFallback = fallback;
            Assert.Same(fallback, propertyMap.Pipeline.EmptyFallback);

            propertyMap.Pipeline.EmptyFallback = null;
            Assert.Null(propertyMap.Pipeline.EmptyFallback);
        }

        [Fact]
        public void InvalidFallback_Set_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo);

            var fallback = new FixedValueFallback(10);
            propertyMap.Pipeline.InvalidFallback = fallback;
            Assert.Same(fallback, propertyMap.Pipeline.InvalidFallback);

            propertyMap.Pipeline.InvalidFallback = null;
            Assert.Null(propertyMap.Pipeline.InvalidFallback);
        }

        [Fact]
        public void AddCellValueMapper_ValidItem_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo);
            var item1 = new BoolMapper();
            var item2 = new BoolMapper();

            propertyMap.Pipeline.AddCellValueMapper(item1);
            propertyMap.Pipeline.AddCellValueMapper(item2);
            Assert.Equal(new ICellValueMapper[] { item1, item2 }, propertyMap.Pipeline.CellValueMappers);
        }

        [Fact]
        public void AddCellValueMapper_NullItem_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("mapper", () => propertyMap.Pipeline.AddCellValueMapper(null));
        }

        [Fact]
        public void RemoveCellValueMapper_Index_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo);
            propertyMap.Pipeline.AddCellValueMapper(new BoolMapper());

            propertyMap.Pipeline.RemoveCellValueMapper(0);
            Assert.Empty(propertyMap.Pipeline.CellValueMappers);
        }

        [Fact]
        public void AddCellValueTransformer_ValidTransformer_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo);
            var transformer1 = new TrimCellValueTransformer();
            var transformer2 = new TrimCellValueTransformer();

            propertyMap.Pipeline.AddCellValueTransformer(transformer1);
            propertyMap.Pipeline.AddCellValueTransformer(transformer2);
            Assert.Equal(new ICellValueTransformer[] { transformer1, transformer2 }, propertyMap.Pipeline.CellValueTransformers);
        }

        [Fact]
        public void AddCellValueTransformer_NullTransformer_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("transformer", () => propertyMap.Pipeline.AddCellValueTransformer(null));
        }

        [Fact]
        public void CellReader_SetValid_GetReturnsExpected()
        {
            var cellReader = new ColumnNameValueReader("ColumnName");
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new OneToOnePropertyMap(propertyInfo)
            {
                CellReader = cellReader
            };

            Assert.Same(cellReader, propertyMap.CellReader);
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
