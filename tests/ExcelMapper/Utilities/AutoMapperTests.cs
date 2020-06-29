using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace ExcelMapper.Utilities.Tests
{
    public class AutoMapperTests
    {
        [Theory]
        [InlineData(FallbackStrategy.ThrowIfPrimitive)]
        [InlineData(FallbackStrategy.SetToDefaultValue)]
        public void TryCreateClassMap_ValidType_ReturnsTrue(FallbackStrategy emptyValueStrategy)
        {
            Assert.True(AutoMapper.TryCreateClassMap<TestClass>(emptyValueStrategy, out ExcelClassMap<TestClass> classMap));
            Assert.Equal(emptyValueStrategy, classMap.EmptyValueStrategy);
            Assert.Equal(typeof(TestClass), classMap.Type);
            Assert.Equal(5, classMap.Properties.Count);

            IEnumerable<string> members = classMap.Properties.Select(m => m.Member.Name);
            Assert.Contains("_inheritedField", members);
            Assert.Contains("_field", members);
            Assert.Contains("InheritedProperty", members);
            Assert.Contains("Property", members);
            Assert.Contains("PrivateGetProperty", members);
        }

        [Fact]
        public void TryCreateClassMap_InterfaceType_ReturnsFalse()
        {
            Assert.False(AutoMapper.TryCreateClassMap<IConvertible>(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IConvertible> classMap));
            Assert.Null(classMap);
        }

        [Theory]
        [InlineData(FallbackStrategy.ThrowIfPrimitive - 1)]
        [InlineData(FallbackStrategy.SetToDefaultValue + 1)]
        public void TryCreateClassMap_InvalidFallbackStrategy_ThrowsArgumentException(FallbackStrategy emptyValueStrategy)
        {
            Assert.Throws<ArgumentException>("emptyValueStrategy", () => AutoMapper.TryCreateClassMap<IConvertible>(emptyValueStrategy, out _));
        }

#pragma warning disable 0649, 0169
        private class BaseClass
        {
            public string _inheritedField;
            public string InheritedProperty { get; set; }
        }

        private class TestClass : BaseClass
        {
            public string _field;
            protected string _protectedField;
            internal string _internalField;
            private string _privateField;
            public static string s_field;

            public string Property { get; set; }
            protected string ProtectedProperty { get; set; }
            internal string InternalProperty { get; set; }
            private string PrivateProperty { get; set; }
            public static string StaticProperty { get; set; }

            public string PrivateSetProperty { get; private set; }
            public string PrivateGetProperty { private get; set; }
        }
#pragma warning restore 0649
    }
}