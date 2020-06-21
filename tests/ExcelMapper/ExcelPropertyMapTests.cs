using System;
using System.Reflection;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class PropertyMapTests
    {
        [Fact]
        public void Ctor_PropertyInfoMember_Success()
        {
            MemberInfo propertyInfo = typeof(ClassWithEvent).GetProperty(nameof(ClassWithEvent.Property));
            var map = new OneToOneMap<int>(new ColumnNameValueReader("Property"));

            var propertyMap = new ExcelPropertyMap(propertyInfo, map);
            Assert.Same(propertyInfo, propertyMap.Member);
            Assert.Same(map, propertyMap.Map);

            var instance = new ClassWithEvent();
            propertyMap.SetValueFactory(instance, 10);
            Assert.Equal(10, instance.Property);
        }

        [Fact]
        public void Ctor_FieldInfoMember_Success()
        {
            MemberInfo fieldInfo = typeof(ClassWithEvent).GetField(nameof(ClassWithEvent._field));
            var map = new OneToOneMap<int>(new ColumnNameValueReader("Property"));

            var propertyMap = new ExcelPropertyMap(fieldInfo, map);
            Assert.Same(fieldInfo, propertyMap.Member);
            Assert.Same(map, propertyMap.Map);

            var instance = new ClassWithEvent();
            propertyMap.SetValueFactory(instance, 10);
            Assert.Equal(10, instance._field);
        }

        [Fact]
        public void Ctor_NullMember_ThrowsArgumentNullException()
        {
            var map = new OneToOneMap<int>(new ColumnNameValueReader("Property"));
            Assert.Throws<ArgumentNullException>("member", () => new ExcelPropertyMap(null, map));
        }

        [Fact]
        public void Ctor_NullMap_ThrowsArgumentNullException()
        {
            MemberInfo member = typeof(ClassWithEvent).GetField(nameof(ClassWithEvent._field));
            Assert.Throws<ArgumentNullException>("map", () => new ExcelPropertyMap(member, null));
        }

        [Fact]
        public void Ctor_MemberNotFieldOrProperty_ThrowsArgumentException()
        {
            MemberInfo eventInfo = typeof(ClassWithEvent).GetEvent(nameof(ClassWithEvent.Event));
            var map = new OneToOneMap<int>(new ColumnNameValueReader("Property"));
            Assert.Throws<ArgumentException>("member", () => new ExcelPropertyMap(eventInfo, map));
        }

        [Fact]
        public void Ctor_PropertyReadOnly_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(ClassWithEvent).GetProperty(nameof(ClassWithEvent.ReadOnlyProperty));
            var map = new OneToOneMap<int>(new ColumnNameValueReader("Property"));
            Assert.Throws<ArgumentException>("member", () => new ExcelPropertyMap(propertyInfo, map));
        }

        private class ClassWithEvent
        {
            public event EventHandler Event { add { } remove { } }

            public int Property { get; set; }
#pragma warning disable 0649
            public int _field;
#pragma warning restore 0649

            public int ReadOnlyProperty { get; }
        }
    }
}
