using System;
using Xunit;

namespace ExcelMapper.Factories;

public class ArrayEnumerableFactoryTests
{
    [Fact]
    public void Begin_End_Success()
    {
        var factory = new ArrayEnumerableFactory<int>();

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([0], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([0], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new ArrayEnumerableFactory<int>();
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }
    [Fact]
    public void Add_End_Success()
    {
        var factory = new ArrayEnumerableFactory<int>();

        // Begin.
        factory.Begin(1);
        factory.Add(1);
        var value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([1], value);

        // Begin again.
        factory.Begin(1);
        factory.Add(2);
        value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([2], value);
    }

    [Fact]
    public void Add_OutOfRange_ThrowsIndexOutOfRangeException()
    {
        var factory = new ArrayEnumerableFactory<int>();
        factory.Begin(1);
        factory.Add(2);

        Assert.Throws<IndexOutOfRangeException>(() => factory.Add(3));
        
        var value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([2], value);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ArrayEnumerableFactory<int>();
        Assert.Throws<ExcelMappingException>(() => factory.Add(1));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ArrayEnumerableFactory<int>();
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new ArrayEnumerableFactory<int>();
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new ArrayEnumerableFactory<int>();
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([0], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new ArrayEnumerableFactory<int>();
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<int[]>(factory.End());
        Assert.Equal([0], value);
    }
}
