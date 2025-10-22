using ExcelMapper.Abstractions;

namespace ExcelMapper.Fallbacks.Tests;

public class FixedValueFallbackTests
{
    [Theory]
    [InlineData(null)]
    [InlineData(1)]
    [InlineData("value")]
    public void Ctor_Object(object? value)
    {
        var fallback = new FixedValueFallback(value);
        Assert.Same(value, fallback.Value);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(1)]
    [InlineData("value")]
    public void PerformFallback_Invoke_ReturnsFixedValue(object? value)
    {
        var fallback = new FixedValueFallback(value);

        var result = fallback.PerformFallback(null!, 0, new ReadCellResult(), null, null!);
        Assert.Same(value, result);
    }
}
