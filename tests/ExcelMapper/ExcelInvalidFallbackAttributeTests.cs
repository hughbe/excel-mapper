using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;

namespace ExcelMapper.Tests;

public class ExcelInvalidFallbackAttributeTests
{
    [Theory]
    [InlineData(typeof(InvalidFallback))]
    [InlineData(typeof(FixedValueFallback))]
    [InlineData(typeof(ThrowFallback))]
    [InlineData(typeof(NoConstructorInvalidFallback))]
    public void Ctor_Type(Type fallbackType)
    {
        var attribute = new ExcelInvalidFallbackAttribute(fallbackType);
        Assert.Same(fallbackType, attribute.Type);
        Assert.Null(attribute.ConstructorArguments);
    }

    [Fact]
    public void Ctor_NullFallbackType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("fallbackType", () => new ExcelInvalidFallbackAttribute(null!));
    }


    [Theory]
    [InlineData(typeof(IFallbackItem))]
    [InlineData(typeof(ISubInvalidFallback))]
    [InlineData(typeof(AbstractInvalidFallback))]
    [InlineData(typeof(int))]
    [InlineData(typeof(object))]
    [InlineData(typeof(ExcelInvalidFallbackAttributeTests))]
    public void Ctor_InvalidFallbackType_ThrowsArgumentException(Type fallbackType)
    {
        Assert.Throws<ArgumentException>("fallbackType", () => new ExcelInvalidFallbackAttribute(fallbackType));
    }

    private interface ISubInvalidFallback : IFallbackItem
    {
    }

    private abstract class AbstractInvalidFallback : IFallbackItem
    {
        public abstract object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member);
    }

    private class NoConstructorInvalidFallback : IFallbackItem
    {
        private NoConstructorInvalidFallback()
        {
        }

        public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
            => throw new NotImplementedException();
    }

    public static IEnumerable<object?[]> ConstructorArguments_Set_TestData()
    {
        yield return new object?[] { null };
        yield return new object[] { new object[0] };
        yield return new object[] { new object?[] { "Value", null } };
    }

    [Theory]
    [MemberData(nameof(ConstructorArguments_Set_TestData))]
    public void ConstructorArguments_Set_GetReturnsExpected(object?[]? value)
    {
        var attribute = new ExcelInvalidFallbackAttribute(typeof(InvalidFallback))
        {
            ConstructorArguments = value
        };
        Assert.Same(value, attribute.ConstructorArguments);
        
        // Set.
        attribute.ConstructorArguments = value;
        Assert.Same(value, attribute.ConstructorArguments);
    }

    private class InvalidFallback : IFallbackItem
    {
        public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
            => throw new NotImplementedException();
    }
}
