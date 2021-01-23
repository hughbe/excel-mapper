using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Fallbacks.Tests
{
    public class ThrowFallbackTests
    {
        public static IEnumerable<object[]> InnerException_TestData()
        {
            yield return new object[] { null };
            yield return new object[] { new DivideByZeroException() };
        }

        [Theory]
        [MemberData(nameof(InnerException_TestData))]
        public void GetProperty_InvokePropertyInfo_ThrowsExcelMappingException(Exception exception)
        {
            var fallback = new ThrowFallback();
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));

            Assert.Throws<ExcelMappingException>(() => fallback.PerformFallback(null, 0, new ReadCellValueResult(), exception, propertyInfo));
        }

        [Theory]
        [MemberData(nameof(InnerException_TestData))]
        public void GetProperty_InvokeFieldInfo_ThrowsExcelMappingException(Exception exception)
        {
            var fallback = new ThrowFallback();
            MemberInfo fieldInfo = typeof(TestClass).GetField(nameof(TestClass._field));

            Assert.Throws<ExcelMappingException>(() => fallback.PerformFallback(null, 0, new ReadCellValueResult(), exception, fieldInfo));
        }

        [Theory]
        [MemberData(nameof(InnerException_TestData))]
        public void GetProperty_InvokeEventInfo_ThrowsArgumentException(Exception exception)
        {
            var fallback = new ThrowFallback();
            MemberInfo eventInfo = typeof(TestClass).GetEvent(nameof(TestClass.Event));

            Assert.Throws<ArgumentException>("member", () => fallback.PerformFallback(null, 0, new ReadCellValueResult(), exception, eventInfo));
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
