using System;
using System.Reflection;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Fallbacks.Tests
{
    public class ThrowFallbackTests
    {
        [Fact]
        public void GetProperty_InvokePropertyInfo_ThrowsExcelMappingException()
        {
            var fallback = new ThrowFallback();
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));

            Assert.Throws<ExcelMappingException>(() => fallback.PerformFallback(null, 0, new ReadCellValueResult(), propertyInfo));
        }

        [Fact]
        public void GetProperty_InvokeFieldInfo_ThrowsExcelMappingException()
        {
            var fallback = new ThrowFallback();
            MemberInfo fieldInfo = typeof(TestClass).GetField(nameof(TestClass._field));

            Assert.Throws<ExcelMappingException>(() => fallback.PerformFallback(null, 0, new ReadCellValueResult(), fieldInfo));
        }

        [Fact]
        public void GetProperty_InvokeEventInfo_ThrowsArgumentException()
        {
            var fallback = new ThrowFallback();
            MemberInfo eventInfo = typeof(TestClass).GetEvent(nameof(TestClass.Event));

            Assert.Throws<ArgumentException>("member", () => fallback.PerformFallback(null, 0, new ReadCellValueResult(), eventInfo));
        }

        private class TestClass
        {
            public string Value { get; set; }
#pragma warning disable 0649
            public string _field;
#pragma warning restore 0649

            public event EventHandler Event { add { } remove { } }
        }
    }
}
