using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Fallbacks.Tests;

public class ThrowFallbackTests
{
    public static IEnumerable<object?[]> InnerException_TestData()
    {
        yield return new object?[] { null };
        yield return new object?[] { new DivideByZeroException() };
    }

    [Theory]
    [MemberData(nameof(InnerException_TestData))]
    public void PerformFallback_InvokePropertyInfo_ThrowsExcelMappingException(Exception? exception)
    {
        var fallback = new ThrowFallback();
        MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;

        Assert.Throws<ExcelMappingException>(() => fallback.PerformFallback(null!, 0, new ReadCellResult(), exception, propertyInfo));
    }

    [Theory]
    [MemberData(nameof(InnerException_TestData))]
    public void PerformFallback_InvokeFieldInfo_ThrowsExcelMappingException(Exception exception)
    {
        var fallback = new ThrowFallback();
        var fieldInfo = typeof(TestClass).GetField(nameof(TestClass._field))!;

        Assert.Throws<ExcelMappingException>(() => fallback.PerformFallback(null!, 0, new ReadCellResult(), exception, fieldInfo));
    }

    [Theory]
    [MemberData(nameof(InnerException_TestData))]
    public void PerformFallback_InvokeEventInfo_ThrowsArgumentException(Exception exception)
    {
        var fallback = new ThrowFallback();
        var eventInfo = typeof(TestClass).GetEvent(nameof(TestClass.Event))!;

        Assert.Throws<ArgumentException>("member", () => fallback.PerformFallback(null!, 0, new ReadCellResult(), exception, eventInfo));
    }

    [Theory]
    [MemberData(nameof(InnerException_TestData))]
    public void PerformFallback_InvokeNullMember_ThrowsExcelMappingException(Exception exception)
    {
        var fallback = new ThrowFallback();
        Assert.Throws<ExcelMappingException>(() => fallback.PerformFallback(null!, 0, new ReadCellResult(), exception, null));
    }

    private class TestClass
    {
        public string Value { get; set; } = default!;
#pragma warning disable 0649
        public string _field = default!;
#pragma warning restore 0649

        public event EventHandler Event { add { } remove { } }
    }
}
