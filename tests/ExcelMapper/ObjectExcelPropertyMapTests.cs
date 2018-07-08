using System;
using System.Reflection;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ObjectExcelPropertyMapTests
    {
        [Fact]
        public void WithClassMap_ClassMapFactory_ReturnsExpected()
        {
            bool calledClassMapFactory = false;
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new ObjectExcelPropertyMap<string>(propertyInfo, new ExcelClassMap<string>());
            Action<ExcelClassMap<string>> classMapFactory = classMap =>
            {
                calledClassMapFactory = true;
                Assert.Same(classMap, propertyMap.ClassMap);
            };

            Assert.Same(propertyMap, propertyMap.WithClassMap(classMapFactory));
            Assert.True(calledClassMapFactory);
        }

        [Fact]
        public void WithClassMap_NullClassMapFactory_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new ObjectExcelPropertyMap<string>(propertyInfo, new ExcelClassMap<string>());

            Assert.Throws<ArgumentNullException>("classMapFactory", () => propertyMap.WithClassMap((Action<ExcelClassMap<string>>)null));
        }

        [Fact]
        public void WithClassMap_ClassMap_ReturnsExpected()
        {
            var classMap = new ExcelClassMap<string>();
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));

            var propertyMap = new ObjectExcelPropertyMap<string>(propertyInfo, new ExcelClassMap<string>());
            Assert.Same(propertyMap, propertyMap.WithClassMap(classMap));
            Assert.Same(classMap, propertyMap.ClassMap);
        }

        [Fact]
        public void WithClassMap_NullClassMap_ThrowsArgumentNullException()
        {
            var classMap = new ExcelClassMap<string>();
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new ObjectExcelPropertyMap<string>(propertyInfo, new ExcelClassMap<string>())
            {
                ClassMap = classMap
            };

            Assert.Throws<ArgumentNullException>("classMap", () => propertyMap.WithClassMap((ExcelClassMap<string>)null));
        }

        [Fact]
        public void ClassMap_SetValid_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new ObjectExcelPropertyMap<string>(propertyInfo, new ExcelClassMap<string>());

            Assert.Throws<ArgumentNullException>("value", () => propertyMap.ClassMap = null);
        }

        [Fact]
        public void ClassMap_SetNull_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new ObjectExcelPropertyMap<string>(propertyInfo, new ExcelClassMap<string>());

            Assert.Throws<ArgumentNullException>("value", () => propertyMap.ClassMap = null);
        }

        private class TestClass
        {
            public string Value { get; set; }
        }
    }
}
