using System;
using System.Reflection;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class PropertyMapCollectionTests
    {
        [Fact]
        public void Add_ValidItem_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property));
            var map1 = new OneToOneMap<int>(new ColumnNameValueReader("Property"));
            var propertyMap1 = new ExcelPropertyMap(propertyInfo, map1);
            var map2 = new OneToOneMap<int>(new ColumnNameValueReader("Property"));
            var propertyMap2 = new ExcelPropertyMap(propertyInfo, map2);
            ExcelPropertyMapCollection mappings = new TestClassMap().Properties;

            mappings.Add(propertyMap1);
            Assert.Same(propertyMap1, Assert.Single(mappings));
            Assert.Same(propertyMap1, mappings[0]);

            mappings.Add(propertyMap2);
            Assert.Equal(2, mappings.Count);
            Assert.Same(propertyMap2, mappings[1]);

            mappings.Add(propertyMap1);
            Assert.Equal(3, mappings.Count);
            Assert.Same(propertyMap1, mappings[2]);
        }

        [Fact]
        public void Add_NullItem_ThrowsArgumentNullException()
        {
            ExcelPropertyMapCollection mappings = new TestClassMap().Properties;
            Assert.Throws<ArgumentNullException>("item", () => mappings.Add(null));
        }

        [Fact]
        public void Insert_ValidItem_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property));
            var map1 = new OneToOneMap<int>(new ColumnNameValueReader("Property"));
            var propertyMap1 = new ExcelPropertyMap(propertyInfo, map1);
            var map2 = new OneToOneMap<int>(new ColumnNameValueReader("Property"));
            var propertyMap2 = new ExcelPropertyMap(propertyInfo, map2);
            ExcelPropertyMapCollection mappings = new TestClassMap().Properties;

            mappings.Insert(0, propertyMap1);
            Assert.Same(propertyMap1, Assert.Single(mappings));
            Assert.Same(propertyMap1, mappings[0]);

            mappings.Insert(0, propertyMap2);
            Assert.Equal(2, mappings.Count);
            Assert.Same(propertyMap2, mappings[0]);

            mappings.Insert(1, propertyMap1);
            Assert.Equal(3, mappings.Count);
            Assert.Same(propertyMap1, mappings[1]);
        }

        [Fact]
        public void Insert_NullItem_ThrowsArgumentNullException()
        {
            ExcelPropertyMapCollection mappings = new TestClassMap().Properties;
            Assert.Throws<ArgumentNullException>("item", () => mappings.Insert(0, null));
        }

        [Fact]
        public void Item_SetValidItem_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property));
            var map1 = new OneToOneMap<int>(new ColumnNameValueReader("Property"));
            var propertyMap1 = new ExcelPropertyMap(propertyInfo, map1);
            var map2 = new OneToOneMap<int>(new ColumnNameValueReader("Property"));
            var propertyMap2 = new ExcelPropertyMap(propertyInfo, map2);
            ExcelPropertyMapCollection mappings = new TestClassMap().Properties;
            mappings.Add(propertyMap1);

            mappings[0] = propertyMap2;
            Assert.Same(propertyMap2, mappings[0]);
        }

        [Fact]
        public void Item_SetNull_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property));
            var map = new OneToOneMap<int>(new ColumnNameValueReader("Property"));
            var propertyMap = new ExcelPropertyMap(propertyInfo, map);
            ExcelPropertyMapCollection mappings = new TestClassMap().Properties;
            mappings.Add(propertyMap);

            Assert.Throws<ArgumentNullException>("item", () => mappings[0] = null);
        }

        private class TestClassMap : ExcelClassMap<Helpers.TestClass>
        {
        }


        private class TestClass
        {
            public int Property { get; set; }
        }
    }
}
