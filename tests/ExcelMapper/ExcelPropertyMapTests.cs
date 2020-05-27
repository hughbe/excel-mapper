using System;
using System.Reflection;
using ExcelDataReader;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelPropertyMapTests
    {
        [Fact]
        public void Ctor_PropertyInfoMember_Success()
        {
            MemberInfo propertyInfo = typeof(ClassWithEvent).GetProperty(nameof(ClassWithEvent.Property));

            var propertyMap = new SubPropertyMap(propertyInfo);
            Assert.Same(propertyInfo, propertyMap.Member);

            var instance = new ClassWithEvent();
            propertyMap.SetPropertyFactory(instance, 10);
            Assert.Equal(10, instance.Property);
        }

        [Fact]
        public void Ctor_FieldInfoMember_Success()
        {
            MemberInfo fieldInfo = typeof(ClassWithEvent).GetField(nameof(ClassWithEvent._field));

            var propertyMap = new SubPropertyMap(fieldInfo);
            Assert.Same(fieldInfo, propertyMap.Member);

            var instance = new ClassWithEvent();
            propertyMap.SetPropertyFactory(instance, 10);
            Assert.Equal(10, instance._field);
        }

        [Fact]
        public void Ctor_NullMember_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("member", () => new SubPropertyMap(null));
        }

        [Fact]
        public void Ctor_MemberNotFieldOrProperty_ThrowsArgumentException()
        {
            MemberInfo eventInfo = typeof(ClassWithEvent).GetEvent(nameof(ClassWithEvent.Event));
            Assert.Throws<ArgumentException>("member", () => new SubPropertyMap(eventInfo));
        }

        [Fact]
        public void Ctor_PropertyReadOnly_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(ClassWithEvent).GetProperty(nameof(ClassWithEvent.ReadOnlyProperty));
            Assert.Throws<ArgumentException>("member", () => new SubPropertyMap(propertyInfo));
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

        private class SubPropertyMap : ExcelPropertyMap
        {
            public SubPropertyMap(MemberInfo member) : base(member) { }

            public override void SetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, object instance)
            {
            }
        }
    }
}
