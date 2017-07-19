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
    public class SingleExcelPropertyMapTests
    {
        [Fact]
        public void Ctor_Member_Type_EmptyValueStrategy()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));

            var propertyMap = new SingleExcelPropertyMap(propertyInfo);
            Assert.Same(propertyInfo, propertyMap.Member);

            Assert.Empty(propertyMap.CellValueMappers);
            Assert.Empty(propertyMap.CellValueTransformers);
        }

        [Fact]
        public void EmptyFallback_Set_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SingleExcelPropertyMap(propertyInfo);

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
            var propertyMap = new SingleExcelPropertyMap(propertyInfo);

            var fallback = new FixedValueFallback(10);
            propertyMap.InvalidFallback = fallback;
            Assert.Same(fallback, propertyMap.InvalidFallback);

            propertyMap.InvalidFallback = null;
            Assert.Null(propertyMap.InvalidFallback);
        }

        [Fact]
        public void AddCellValueMapper_ValidItem_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SingleExcelPropertyMap(propertyInfo);
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
            var propertyMap = new SingleExcelPropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("mapper", () => propertyMap.AddCellValueMapper(null));
        }

        [Fact]
        public void RemoveCellValueMapper_Index_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SingleExcelPropertyMap(propertyInfo);
            propertyMap.AddCellValueMapper(new BoolMapper());

            propertyMap.RemoveCellValueMapper(0);
            Assert.Empty(propertyMap.CellValueMappers);
        }

        [Fact]
        public void AddCellValueTransformer_ValidTransformer_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SingleExcelPropertyMap(propertyInfo);
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
            var propertyMap = new SingleExcelPropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("transformer", () => propertyMap.AddCellValueTransformer(null));
        }

        [Fact]
        public void CellReader_SetValid_GetReturnsExpected()
        {
            var cellReader = new ColumnNameValueReader("ColumnName");
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SingleExcelPropertyMap(propertyInfo)
            {
                CellReader = cellReader
            };

            Assert.Same(cellReader, propertyMap.CellReader);
        }

        [Fact]
        public void CellReader_SetNull_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SingleExcelPropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("value", () => propertyMap.CellReader = null);
        }

        private class TestClass
        {
            public string Value { get; set; }
        }
    }
}
