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

        private class TestClassMap : ExcelClassMap<Helpers.TestClass>
        {
            public TestClassMap(EmptyValueStrategy emptyValueStrategy) : base(emptyValueStrategy) { }
        }
    }
}
