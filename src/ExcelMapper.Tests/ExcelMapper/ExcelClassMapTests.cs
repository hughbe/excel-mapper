using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelClassMapTests : ExcelClassMap<Helpers.TestClass>
    {
        [Fact]
        public void Map_ExpressionNotMemberExpression_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("expression", () => Map(p => new List<string>()));
        }

        [Fact]
        public void Map_IEnumerable_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => Map(p => p.ConcreteIEnumerable));
        }

        [Fact]
        public void MultiMap_UnknownInterface_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => Map<string>(p => p.UnknownInterfaceValue));
        }

        [Fact]
        public void MultiMap_ConcreteIEnumerable_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => Map<string>(p => p.ConcreteIEnumerable));
        }

        [Fact]
        public void MultiMap_CantMapIEnumerableElementType_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => Map(p => p.CantMapElementType));
        }

        [Fact]
        public void MapObject_Interface_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => MapObject(p => p.UnknownInterfaceValue));
        }

        [Fact]
        public void MapObject_InvalidMemberType_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => MapObject(p => p.InvalidMemberType));
        }

        [Fact]
        public void Map_InvalidTargetType_ThrowsArgumentException()
        {
            var otherType = new OtherType();
            Assert.Throws<ArgumentException>("expression", () => Map(p => otherType.Value));
        }

        public class OtherType
        {
            public string Value { get; set; }
        }

        [Theory]
        [InlineData(EmptyValueStrategy.ThrowIfPrimitive)]
        [InlineData(EmptyValueStrategy.SetToDefaultValue)]
        public void Ctor_EmptyValueStrategy(EmptyValueStrategy emptyValueStrategy)
        {
            var map = new TestClassMap(emptyValueStrategy);
            Assert.Equal(emptyValueStrategy, map.EmptyValueStrategy);
            Assert.Equal(typeof(Helpers.TestClass), map.Type);
        }

        [Theory]
        [InlineData(EmptyValueStrategy.ThrowIfPrimitive - 1)]
        [InlineData(EmptyValueStrategy.SetToDefaultValue + 1)]
        public void Ctor_InvalidEmptyValueStrategy_ThrowsArgumentException(EmptyValueStrategy emptyValueStrategy)
        {
            Assert.Throws<ArgumentException>("emptyValueStrategy", () => new TestClassMap(emptyValueStrategy));
        }

        [Fact]
        public void MapObject_ClassMapFactory_ReturnsExpected()
        {
            var map = new TestClassMap(EmptyValueStrategy.ThrowIfPrimitive);
            ObjectPropertyMapping<string> mapping = map.MapObject(t => t.Value);
            Assert.NotNull(mapping.ClassMap);
        }

        private class TestClassMap : ExcelClassMap<Helpers.TestClass>
        {
            public TestClassMap(EmptyValueStrategy emptyValueStrategy) : base(emptyValueStrategy) { }
        }
    }
}
