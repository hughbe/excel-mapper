namespace ExcelMapper.Tests;

public class ExcelEmptyFallbackAttributeTests
{
    [Theory]
    [InlineData(typeof(EmptyFallback))]
    [InlineData(typeof(NoConstructorEmptyFallback))]
    [InlineData(typeof(MultipleFallback))]
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
    [InlineData(typeof(ICellFallback<int>))]
    [InlineData(typeof(ISubEmptyFallback))]
    [InlineData(typeof(AbstractEmptyFallback))]
    [InlineData(typeof(int))]
    [InlineData(typeof(object))]
    [InlineData(typeof(ICellTransformer))]
    [InlineData(typeof(ExcelEmptyFallbackAttributeTests))]
    public void Ctor_InvalidFallbackType_ThrowsArgumentException(Type fallbackType)
    {
        Assert.Throws<ArgumentException>("fallbackType", () => new ExcelEmptyFallbackAttribute(fallbackType));
    }

    private interface ISubEmptyFallback : ICellFallback<int>
    {
    }

    private abstract class AbstractEmptyFallback : ICellFallback<int>
    {
        public abstract int Fallback(CellValue value, CellContext context, Exception? exception);
    }

    private class NoConstructorEmptyFallback : ICellFallback<int>
    {
        private NoConstructorEmptyFallback()
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
        var attribute = new ExcelEmptyFallbackAttribute(typeof(EmptyFallback))
        {
            ConstructorArguments = value
        };
        Assert.Same(value, attribute.ConstructorArguments);
        
        // Set.
        attribute.ConstructorArguments = value;
        Assert.Same(value, attribute.ConstructorArguments);
    }

    private sealed class EmptyFallback : ICellFallback<int>
    {
        public int Fallback(CellValue value, CellContext context, Exception? exception) => 1;
    }

    private sealed class MultipleFallback : ICellFallback<int>, ICellFallback<string>
    {
        int ICellFallback<int>.Fallback(CellValue value, CellContext context, Exception? exception) => 1;

        string ICellFallback<string>.Fallback(CellValue value, CellContext context, Exception? exception) => "1";
    }
}
