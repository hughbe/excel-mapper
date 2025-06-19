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
            Assert.True(AutoMapper.TryCreateClassMap<TestClass>(emptyValueStrategy, out ExcelClassMap<TestClass>? classMap));
            Assert.Equal(emptyValueStrategy, classMap!.EmptyValueStrategy);
            Assert.Equal(typeof(TestClass), classMap.Type);
            Assert.Equal(5, classMap.Properties.Count);

            List<string> members = classMap.Properties.Select(m => m.Member.Name).ToList();
            Assert.Contains("_inheritedField", members);
            Assert.Contains("_field", members);
            Assert.Contains("InheritedProperty", members);
            Assert.Contains("Property", members);
            Assert.Contains("PrivateGetProperty", members);
        }

        [Fact]
        public void TryCreateClassMap_InterfaceType_ReturnsFalse()
        {
            Assert.False(AutoMapper.TryCreateClassMap<IConvertible>(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IConvertible>? classMap));
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
            public string _inheritedField = default!;
            public string InheritedProperty { get; set; } = default!;
        }

        private class TestClass : BaseClass
        {
            public string _field = default!;
            protected string _protectedField = default!;
            internal string _internalField = default!;
#pragma warning disable CS0414 // Field is never assigned to, and will always have its default value
            private string _privateField = default!;
#pragma warning restore CS0414 // Field is never assigned to, and will always have its default value
            public static string s_field = default!;
            
            public string Property { get; set; } = default!;
            protected string ProtectedProperty { get; set; } = default!;
            internal string InternalProperty { get; set; } = default!;
            private string PrivateProperty { get; set; } = default!;
            public static string StaticProperty { get; set; } = default!;

            public string PrivateSetProperty { get; private set; } = default!;
            public string PrivateGetProperty { private get; set; } = default!;
        }
#pragma warning restore 0649
    }
}