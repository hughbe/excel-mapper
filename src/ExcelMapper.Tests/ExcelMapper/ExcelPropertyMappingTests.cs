using System;
using System.Reflection;
using ExcelDataReader;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelPropertyMappingTests
    {
        [Fact]
        public void Ctor_PropertyInfoMember_Success()
        {
            MemberInfo propertyInfo = typeof(ClassWithEvent).GetProperty(nameof(ClassWithEvent.Property));

            var mapping = new SubPropertyMapping(propertyInfo);
            Assert.Same(propertyInfo, mapping.Member);

            var instance = new ClassWithEvent();
            mapping.SetPropertyFactory(instance, 10);
            Assert.Equal(10, instance.Property);
        }

        [Fact]
        public void Ctor_FieldInfoMember_Success()
        {
            MemberInfo fieldInfo = typeof(ClassWithEvent).GetField(nameof(ClassWithEvent._field));

            var mapping = new SubPropertyMapping(fieldInfo);
            Assert.Same(fieldInfo, mapping.Member);

            var instance = new ClassWithEvent();
            mapping.SetPropertyFactory(instance, 10);
            Assert.Equal(10, instance._field);
        }

        [Fact]
        public void Ctor_NullMember_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("member", () => new SubPropertyMapping(null));
        }

        [Fact]
        public void Ctor_MemberNotFieldOrProperty_ThrowsArgumentException()
        {
            MemberInfo eventInfo = typeof(ClassWithEvent).GetEvent(nameof(ClassWithEvent.Event));
            Assert.Throws<ArgumentException>("member", () => new SubPropertyMapping(eventInfo));
        }

        [Fact]
        public void Ctor_PropertyReadOnly_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(ClassWithEvent).GetProperty(nameof(ClassWithEvent.ReadOnlyProperty));
            Assert.Throws<ArgumentException>("member", () => new SubPropertyMapping(propertyInfo));
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

        private class SubPropertyMapping : PropertyMapping
        {
            public SubPropertyMapping(MemberInfo member) : base(member) { }

            public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
            {
                return 10;
            }
        }
    }
}
