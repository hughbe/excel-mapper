namespace ExcelMapper.Tests;

public class ExcelInvalidFallbackAttributeTests
{
    [Theory]
    [InlineData(typeof(InvalidFallback))]
    [InlineData(typeof(NoConstructorInvalidFallback))]
    [InlineData(typeof(MultipleFallback))]
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
    [InlineData(typeof(ICellFallback<int>))]
    [InlineData(typeof(ISubInvalidFallback))]
    [InlineData(typeof(AbstractInvalidFallback))]
    [InlineData(typeof(int))]
    [InlineData(typeof(object))]
    [InlineData(typeof(ICellTransformer))]
    [InlineData(typeof(ExcelInvalidFallbackAttributeTests))]
    public void Ctor_InvalidFallbackType_ThrowsArgumentException(Type fallbackType)
    {
        Assert.Throws<ArgumentException>("fallbackType", () => new ExcelInvalidFallbackAttribute(fallbackType));
    }

    private interface ISubInvalidFallback : ICellFallback<int>
    {
    }

    private abstract class AbstractInvalidFallback : ICellFallback<int>
    {
        public abstract int Fallback(CellValue value, CellContext context, Exception? exception);
    }

    private class NoConstructorInvalidFallback : ICellFallback<int>
    {
        private NoConstructorInvalidFallback()
        {
        }

        public int Fallback(CellValue value, CellContext context, Exception? exception) => throw new NotImplementedException();
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
        var attribute = new ExcelInvalidFallbackAttribute(typeof(InvalidFallback))
        {
            ConstructorArguments = value
        };
        Assert.Same(value, attribute.ConstructorArguments);
        
        // Set.
        attribute.ConstructorArguments = value;
        Assert.Same(value, attribute.ConstructorArguments);
    }

    private sealed class InvalidFallback : ICellFallback<int>
    {
        public int Fallback(CellValue value, CellContext context, Exception? exception) => 1;
    }

    private sealed class MultipleFallback : ICellFallback<int>, ICellFallback<string>
    {
        int ICellFallback<int>.Fallback(CellValue value, CellContext context, Exception? exception) => 1;

        string ICellFallback<string>.Fallback(CellValue value, CellContext context, Exception? exception) => "1";
    }
}
