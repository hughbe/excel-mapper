using System;
using System.Reflection;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ObjectPropertyMappingTests
    {
        [Fact]
        public void WithClassMap_ClassMapFactory_ReturnsExpected()
        {
            bool calledClassMapFactory = false;
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new ObjectPropertyMapping<string>(propertyInfo, new ExcelClassMap<string>());
            Action<ExcelClassMap<string>> classMapFactory = classMap =>
            {
                calledClassMapFactory = true;
                Assert.Same(classMap, mapping.ClassMap);
            };

            Assert.Same(mapping, mapping.WithClassMap(classMapFactory));
            Assert.True(calledClassMapFactory);
        }

        [Fact]
        public void WithClassMap_NullClassMapFactory_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new ObjectPropertyMapping<string>(propertyInfo, new ExcelClassMap<string>());

            Assert.Throws<ArgumentNullException>("classMapFactory", () => mapping.WithClassMap((Action<ExcelClassMap<string>>)null));
        }

        [Fact]
        public void WithClassMap_ClassMap_ReturnsExpected()
        {
            var classMap = new ExcelClassMap<string>();
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));

            var mapping = new ObjectPropertyMapping<string>(propertyInfo, new ExcelClassMap<string>());
            Assert.Same(mapping, mapping.WithClassMap(classMap));
            Assert.Same(classMap, mapping.ClassMap);
        }

        [Fact]
        public void WithClassMap_NullClassMap_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new ObjectPropertyMapping<string>(propertyInfo, new ExcelClassMap<string>());

            Assert.Throws<ArgumentNullException>("classMap", () => mapping.WithClassMap((ExcelClassMap<string>)null));
        }

        private class TestClass
        {
            public string Value { get; set; }
        }
    }
}
