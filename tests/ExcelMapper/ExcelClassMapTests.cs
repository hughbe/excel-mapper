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

            public TypeCode GetTypeCode()
            {
                throw new NotImplementedException();
            }

            public bool ToBoolean(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public byte ToByte(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public char ToChar(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public DateTime ToDateTime(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public decimal ToDecimal(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public double ToDouble(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public short ToInt16(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public int ToInt32(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public long ToInt64(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public sbyte ToSByte(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public float ToSingle(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public string ToString(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public object ToType(Type conversionType, IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public ushort ToUInt16(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public uint ToUInt32(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }

            public ulong ToUInt64(IFormatProvider provider)
            {
                throw new NotImplementedException();
            }
        }

        public class OtherType
        {
            public string Value { get; set; }
        }

        [Theory]
        [InlineData(FallbackStrategy.ThrowIfPrimitive)]
        [InlineData(FallbackStrategy.SetToDefaultValue)]
        public void Ctor_EmptyValueStrategy(FallbackStrategy emptyValueStrategy)
        {
            var map = new TestClassMap(emptyValueStrategy);
            Assert.Equal(emptyValueStrategy, map.EmptyValueStrategy);
            Assert.Equal(typeof(Helpers.TestClass), map.Type);
            Assert.Empty(map.Mappings);
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
            ObjectExcelPropertyMap<string> mapping = map.MapObject(t => t.Value);
            Assert.NotNull(mapping.ClassMap);
        }

        private class TestClassMap : ExcelClassMap<Helpers.TestClass>
        {
            public TestClassMap(FallbackStrategy emptyValueStrategy) : base(emptyValueStrategy) { }
        }
    }
}
