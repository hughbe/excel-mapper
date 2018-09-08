using System;
using System.Reflection;
using ExcelDataReader;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelPropertyMapCollectionTests
    {
        [Fact]
        public void Add_ValidItem_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property));
            var map1 = new SubPropertyMap(propertyInfo);
            var map2 = new SubPropertyMap(propertyInfo);
            ExcelPropertyMapCollection mappings = new TestClassMap().Mappings;
            
            mappings.Add(map1);
            Assert.Same(map1, Assert.Single(mappings));
            Assert.Same(map1, mappings[0]);

            mappings.Add(map2);
            Assert.Equal(2, mappings.Count);
            Assert.Same(map2, mappings[1]);

            mappings.Add(map1);
            Assert.Equal(3, mappings.Count);
            Assert.Same(map1, mappings[2]);
        }

        [Fact]
        public void Add_NullItem_ThrowsArgumentNullException()
        {
            ExcelPropertyMapCollection mappings = new TestClassMap().Mappings;
            Assert.Throws<ArgumentNullException>("item", () => mappings.Add(null));
        }

        [Fact]
        public void Insert_ValidItem_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property));
            var map1 = new SubPropertyMap(propertyInfo);
            var map2 = new SubPropertyMap(propertyInfo);
            ExcelPropertyMapCollection mappings = new TestClassMap().Mappings;
            
            mappings.Insert(0, map1);
            Assert.Same(map1, Assert.Single(mappings));
            Assert.Same(map1, mappings[0]);

            mappings.Insert(0, map2);
            Assert.Equal(2, mappings.Count);
            Assert.Same(map2, mappings[0]);

            mappings.Insert(1, map1);
            Assert.Equal(3, mappings.Count);
            Assert.Same(map1, mappings[1]);
        }

        [Fact]
        public void Insert_NullItem_ThrowsArgumentNullException()
        {
            ExcelPropertyMapCollection mappings = new TestClassMap().Mappings;
            Assert.Throws<ArgumentNullException>("item", () => mappings.Insert(0, null));
        }

        [Fact]
        public void Item_SetValidItem_GetReturnsExpected()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property));
            var map1 = new SubPropertyMap(propertyInfo);
            var map2 = new SubPropertyMap(propertyInfo);
            ExcelPropertyMapCollection mappings = new TestClassMap().Mappings;
            mappings.Add(map1);

            mappings[0] = map2;
            Assert.Same(map2, mappings[0]);
        }

        [Fact]
        public void Item_SetNull_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property));
            var map = new SubPropertyMap(propertyInfo);
            ExcelPropertyMapCollection mappings = new TestClassMap().Mappings;
            mappings.Add(map);

            Assert.Throws<ArgumentNullException>("item", () => mappings[0] = null);
        }

        private class TestClassMap : ExcelClassMap<Helpers.TestClass>
        {
        }


        private class TestClass
        {
            public int Property { get; set; }
        }

        private class SubPropertyMap : ExcelPropertyMap
        {
            public SubPropertyMap(MemberInfo member) : base(member) { }

            public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
            {
                return 10;
            }
        }
    }
}
