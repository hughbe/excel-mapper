using System;
using System.Reflection;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ManyToOneObjectPropertyMapTests
    {
        [Fact]
        public void Ctor_MemberInfo_ExcelClassMap()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var classMap = new ExcelClassMap<string>();
            var propertyMap = new ManyToOneObjectPropertyMap<string>(propertyInfo, classMap);
            Assert.Equal(propertyInfo, propertyMap.Member);
            Assert.Equal(classMap, propertyMap.ClassMap);
        }

        [Fact]
        public void Ctor_NullClassMap_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            Assert.Throws<ArgumentNullException>("classMap", () => new ManyToOneObjectPropertyMap<string>(propertyInfo, null));
        }

        [Fact]
        public void WithClassMap_ClassMapFactory_ReturnsExpected()
        {
            bool calledClassMapFactory = false;
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new ManyToOneObjectPropertyMap<string>(propertyInfo, new ExcelClassMap<string>());
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
            var propertyMap = new ManyToOneObjectPropertyMap<string>(propertyInfo, new ExcelClassMap<string>());

            Assert.Throws<ArgumentNullException>("classMapFactory", () => propertyMap.WithClassMap((Action<ExcelClassMap<string>>)null));
        }

        [Fact]
        public void WithClassMap_ClassMap_ReturnsExpected()
        {
            var classMap = new ExcelClassMap<string>();
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));

            var propertyMap = new ManyToOneObjectPropertyMap<string>(propertyInfo, new ExcelClassMap<string>());
            Assert.Same(propertyMap, propertyMap.WithClassMap(classMap));
            Assert.Same(classMap, propertyMap.ClassMap);
        }

        [Fact]
        public void WithClassMap_NullClassMap_ThrowsArgumentNullException()
        {
            var classMap = new ExcelClassMap<string>();
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new ManyToOneObjectPropertyMap<string>(propertyInfo, new ExcelClassMap<string>())
            {
                ClassMap = classMap
            };

            Assert.Throws<ArgumentNullException>("classMap", () => propertyMap.WithClassMap((ExcelClassMap<string>)null));
        }

        [Fact]
        public void ClassMap_SetValid_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new ManyToOneObjectPropertyMap<string>(propertyInfo, new ExcelClassMap<string>());

            Assert.Throws<ArgumentNullException>("value", () => propertyMap.ClassMap = null);
        }

        [Fact]
        public void ClassMap_SetNull_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new ManyToOneObjectPropertyMap<string>(propertyInfo, new ExcelClassMap<string>());

            Assert.Throws<ArgumentNullException>("value", () => propertyMap.ClassMap = null);
        }

        private class TestClass
        {
            public string Value { get; set; }
        }
    }
}
