using ExcelMapper.Abstractions;

namespace ExcelMapper.Fallbacks.Tests;

public class FixedValueFallbackFactoryTests
{
    [Fact]
    public void Ctor_FuncObject()
    {
        Func<object> factory = () => "test";
        var fallback = new FixedValueFallbackFactory(factory);
        Assert.Same(factory, fallback.Factory);
    }

    [Fact]
    public void Ctor_NullFactory_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("factory", () => new FixedValueFallbackFactory(null!));
    }

    [Theory]
    [InlineData(null)]
    [InlineData(1)]
    [InlineData("value")]
    public void PerformFallback_Invoke_ReturnsFixedValue(object? value)
    {
        var fallback = new FixedValueFallbackFactory(() => value);

        var result = fallback.PerformFallback(null!, 0, new ReadCellResult(), null, null!);
        Assert.Same(value, result);
    }

    [Fact]
    public void PerformFallback_InvokeMultipleTimes_CallsFactoryEachTime()
    {
        int callCount = 0;
        object factory()
        {
            callCount++;
            return callCount;
        }

        var fallback = new FixedValueFallbackFactory(factory);

        var result1 = fallback.PerformFallback(null!, 0, new ReadCellResult(), null, null!);
        var result2 = fallback.PerformFallback(null!, 0, new ReadCellResult(), null, null!);
        var result3 = fallback.PerformFallback(null!, 0, new ReadCellResult(), null, null!);

        Assert.Equal(1, result1);
        Assert.Equal(2, result2);
        Assert.Equal(3, result3);
        Assert.Equal(3, callCount);
    }
}
