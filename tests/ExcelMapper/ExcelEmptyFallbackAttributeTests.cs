using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;

namespace ExcelMapper.Tests;

public class ExcelEmptyFallbackAttributeTests
{
    [Theory]
    [InlineData(typeof(EmptyFallback))]
    [InlineData(typeof(FixedValueFallback))]
    [InlineData(typeof(ThrowFallback))]
    [InlineData(typeof(NoConstructorEmptyFallback))]
    public void Ctor_Type(Type fallbackType)
    {
        var attribute = new ExcelEmptyFallbackAttribute(fallbackType);
        Assert.Same(fallbackType, attribute.Type);
        Assert.Null(attribute.ConstructorArguments);
    }

    [Fact]
    public void Ctor_NullFallbackType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("fallbackType", () => new ExcelEmptyFallbackAttribute(null!));
    }


    [Theory]
    [InlineData(typeof(IFallbackItem))]
    [InlineData(typeof(ISubEmptyFallback))]
    [InlineData(typeof(AbstractEmptyFallback))]
    [InlineData(typeof(int))]
    [InlineData(typeof(object))]
    [InlineData(typeof(ExcelEmptyFallbackAttributeTests))]
    public void Ctor_InvalidFallbackType_ThrowsArgumentException(Type fallbackType)
    {
        Assert.Throws<ArgumentException>("fallbackType", () => new ExcelEmptyFallbackAttribute(fallbackType));
    }

    private interface ISubEmptyFallback : IFallbackItem
    {
    }

    private abstract class AbstractEmptyFallback : IFallbackItem
    {
        public abstract object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member);
    }

    private class NoConstructorEmptyFallback : IFallbackItem
    {
        private NoConstructorEmptyFallback()
        {
        }

        public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
            => throw new NotImplementedException();
    }

    public static IEnumerable<object?[]> ConstructorArguments_Set_TestData()
    {
        yield return new object?[] { null };
        yield return new object[] { Array.Empty<object>() };
        yield return new object[] { new object?[] { "Value", null } };
    }

    [Theory]
    [MemberData(nameof(ConstructorArguments_Set_TestData))]
    public void ConstructorArguments_Set_GetReturnsExpected(object?[]? value)
    {
        var attribute = new ExcelEmptyFallbackAttribute(typeof(EmptyFallback))
        {
            ConstructorArguments = value
        };
        Assert.Same(value, attribute.ConstructorArguments);
        
        // Set.
        attribute.ConstructorArguments = value;
        Assert.Same(value, attribute.ConstructorArguments);
    }

    private class EmptyFallback : IFallbackItem
    {
        public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
            => throw new NotImplementedException();
    }
}
