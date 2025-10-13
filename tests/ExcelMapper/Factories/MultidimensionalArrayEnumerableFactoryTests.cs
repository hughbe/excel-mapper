using System;
using Xunit;

namespace ExcelMapper.Factories;

public class MultidimensionalArrayFactoryTests
{
    [Fact]
    public void Begin_End_Success()
    {
        var factory = new MultidimensionalArrayFactory<int>();

        // Begin.
        factory.Begin([1]);
        var value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([0], value);

        // Begin again.
        factory.Begin([1]);
        value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([0], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        factory.Begin([1]);
        Assert.Throws<ExcelMappingException>(() => factory.Begin([1]));
    }

    [Fact]
    public void Begin_NullLengths_ThrowsArgumentNullException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        Assert.Throws<ArgumentNullException>("lengths", () => factory.Begin(null!));
    }

    [Fact]
    public void Begin_EmptyLengths_ThrowsArgumentException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        Assert.Throws<ArgumentException>("lengths", () => factory.Begin([]));
    }

    [Fact]
    public void Begin_NegativeLength_ThrowsArgumentOutOfRangeException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        Assert.Throws<ArgumentOutOfRangeException>("lengths", () => factory.Begin([1, -1]));
    }

    [Fact]
    public void Set_Invoke_Success()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        factory.Begin([1]);

        factory.Set([0], 1);
        Assert.Equal([1], Assert.IsType<int[]>(factory.End()));
    }

    [Fact]
    public void Set_InvokeOutOfRange_Success()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        factory.Begin([1]);

        factory.Set([0], 1);
        Assert.Throws<IndexOutOfRangeException>(() => factory.Set([1], 1));
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        Assert.Throws<ExcelMappingException>(() => factory.Set([0], 1));
    }

    [Fact]
    public void Set_NullIndices_ThrowsArgumentNullException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        factory.Begin([1]);
        Assert.Throws<ArgumentNullException>("indices", () => factory.Set(null!, 1));
    }

    [Fact]
    public void Set_EmptyIndices_ThrowsArgumentException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        factory.Begin([1]);
        Assert.Throws<ArgumentException>("indices", () => factory.Set([], 1));
    }

    [Fact]
    public void Set_LargeIndices_ThrowsArgumentException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        factory.Begin([1]);
        Assert.Throws<ArgumentException>(null, () => factory.Set([1, 2], 1));
    }

    [Fact]
    public void Set_NegativeIndex_ThrowsArgumentOutOfRangeException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        Assert.Throws<ArgumentOutOfRangeException>("indices", () => factory.Set([-1], 1));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        factory.Begin([1]);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        factory.Begin([1]);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin([1]);
        var value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([0], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new MultidimensionalArrayFactory<int>();
        factory.Reset();

        // Make sure we can begin.
        factory.Begin([1]);
        var value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([0], value);
    }
}
