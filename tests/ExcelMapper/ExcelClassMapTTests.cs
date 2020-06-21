using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelClassMapTTests : ExcelClassMap<Helpers.TestClass>
    {
        [Fact]
        public void Ctor_Default()
        {
            var map = new ExcelClassMap<string>();
            Assert.Equal(typeof(string), map.Type);
            Assert.Empty(map.Properties);
            Assert.Same(map.Properties, map.Properties);
        }

        [Fact]
        public void Map_SingleExpression_Success()
        {
            Map(p => p.Value);
        }

        [Fact]
        public void Map_NestedExpression_Success()
        {
            Map(p => p.NestedValue.IntValue);
        }

        [Fact]
        public void Map_CastExpression_Success()
        {
            Map(p => (string)p.Value);
        }

        [Fact]
        public void Map_NestedCastExpression_Success()
        {
            Map(p => (int)p.NestedValue.IntValue);
        }

        [Fact]
        public void Map_ExpressionNotMemberExpression_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("expression", () => Map(p => new List<string>()));
        }

        [Fact]
        public void Map_NotEnum_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("TProperty", () => Map(p => p.DateValue, ignoreCase: true));
            Assert.Throws<ArgumentException>("TProperty", () => Map(p => p.NullableDateValue, ignoreCase: true));
        }

        [Fact]
        public void Map_IEnumerable_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => Map(p => p.ConcreteIEnumerable));
        }

        [Fact]
        public void Map_IDictionary_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => Map(p => p.ConcreteIDictionary));
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
        public void MultiMap_CantMapIDictionaryValueType_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => Map(p => p.CantMapDictionaryValueType));
        }

        [Fact]
        public void MapObject_Interface_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => MapObject(p => p.UnknownInterfaceValue));
        }

        [Fact]
        public void MapObject_InvalidIListMemberType_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => MapObject(p => p.InvalidIListMemberType));
        }

        [Fact]
        public void MapObject_InvalidIDictionaryMemberType_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => MapObject(p => p.InvalidIDictionaryMemberType));
        }

        [Fact]
        public void Map_InvalidTargetType_ThrowsArgumentException()
        {
            var otherType = new OtherType();
            Assert.Throws<ArgumentException>("expression", () => Map(p => otherType.Value));
        }

        [Fact]
        public void Map_InvalidUnaryExpression_ThrowsArgumentException()
        {
            var otherType = new OtherType();
            Assert.Throws<ArgumentException>("expression", () => Map(p => -otherType.Value));
        }

        [Fact]
        public void Map_InvalidBinaryExpression_ThrowsArgumentException()
        {
            var otherType = new OtherType();
            Assert.Throws<ArgumentException>("expression", () => Map(p => otherType.Value + 1));
        }

        [Fact]
        public void Map_InvalidMethodExpression_ThrowsArgumentException()
        {
            var otherType = new OtherType();
            Assert.Throws<ArgumentException>("expression", () => Map(p => otherType.Value.ToString()));
        }

        [Fact]
        public void Map_InvalidCastExpression_ThrowsExcelMappingException()
        {
            Assert.Throws<ExcelMappingException>(() => Map(p => (CollectionAttribute)p.ObjectValue));
        }

        [Fact]
        public void Map_MultipleMemberAccessTypeAlreadyMapped_ThrowsInvalidOperationException()
        {
            var iconvertibleType = new IConvertibleType();
            var classMap = new ExcelClassMap<IConvertibleValue>();
            classMap.Map(p => p.IConvertibleType);

            Assert.Throws<InvalidOperationException>(() => classMap.Map(p => p.IConvertibleType.Value));
        }

        public class IConvertibleValue
        {
            public IConvertibleType IConvertibleType { get; set; }
        }

        public class IConvertibleType : IConvertible
        {
            public string Value { get; set; }

            public TypeCode GetTypeCode() => throw new NotImplementedException();

            public bool ToBoolean(IFormatProvider provider) => throw new NotImplementedException();

            public byte ToByte(IFormatProvider provider) => throw new NotImplementedException();

            public char ToChar(IFormatProvider provider) => throw new NotImplementedException();

            public DateTime ToDateTime(IFormatProvider provider) => throw new NotImplementedException();

            public decimal ToDecimal(IFormatProvider provider) => throw new NotImplementedException();

            public double ToDouble(IFormatProvider provider) => throw new NotImplementedException();

            public short ToInt16(IFormatProvider provider) => throw new NotImplementedException();

            public int ToInt32(IFormatProvider provider) => throw new NotImplementedException();

            public long ToInt64(IFormatProvider provider) => throw new NotImplementedException();

            public sbyte ToSByte(IFormatProvider provider) => throw new NotImplementedException();

            public float ToSingle(IFormatProvider provider) => throw new NotImplementedException();

            public string ToString(IFormatProvider provider) => throw new NotImplementedException();

            public object ToType(Type conversionType, IFormatProvider provider) => throw new NotImplementedException();

            public ushort ToUInt16(IFormatProvider provider) => throw new NotImplementedException();

            public uint ToUInt32(IFormatProvider provider) => throw new NotImplementedException();

            public ulong ToUInt64(IFormatProvider provider) => throw new NotImplementedException();
        }

        public class OtherType
        {
            public int Value { get; set; }
        }

        [Theory]
        [InlineData(FallbackStrategy.ThrowIfPrimitive)]
        [InlineData(FallbackStrategy.SetToDefaultValue)]
        public void Ctor_EmptyValueStrategy(FallbackStrategy emptyValueStrategy)
        {
            var map = new TestClassMap(emptyValueStrategy);
            Assert.Equal(emptyValueStrategy, map.EmptyValueStrategy);
            Assert.Equal(typeof(Helpers.TestClass), map.Type);
            Assert.Empty(map.Properties);
        }

        [Theory]
        [InlineData(FallbackStrategy.ThrowIfPrimitive - 1)]
        [InlineData(FallbackStrategy.SetToDefaultValue + 1)]
        public void Ctor_InvalidEmptyValueStrategy_ThrowsArgumentException(FallbackStrategy emptyValueStrategy)
        {
            Assert.Throws<ArgumentException>("emptyValueStrategy", () => new TestClassMap(emptyValueStrategy));
        }

        [Fact]
        public void MapObject_ClassMapFactory_ReturnsExpected()
        {
            var map = new TestClassMap(FallbackStrategy.ThrowIfPrimitive);
            ExcelClassMap<string> mapping = map.MapObject(t => t.Value);
            Assert.Empty(mapping.Properties);
        }

        [Fact]
        public void WithClassMap_InvokeClassMapFactory_Success()
        {
            Assert.Same(this, WithClassMap(c =>
            {
                Assert.Same(this, c);
            }));
        }

        [Fact]
        public void WithClassMap_NullClassMapFactory_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("classMapFactory", () => WithClassMap((Action<ExcelClassMap<Helpers.TestClass>>)null));
        }

        [Fact]
        public void WithClassMap_InvokeClassMap_Success()
        {
            var map = new ExcelClassMap<Helpers.TestClass>();
            Assert.Same(this, WithClassMap(map));
        }

        [Fact]
        public void WithClassMap_NullClassMap_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("classMap", () => WithClassMap((ExcelClassMap<Helpers.TestClass>)null));
        }

        private class TestClassMap : ExcelClassMap<Helpers.TestClass>
        {
            public TestClassMap(FallbackStrategy emptyValueStrategy) : base(emptyValueStrategy) { }
        }
    }
}
