using System;
using System.Collections.Frozen;
using Xunit;

namespace ExcelMapper.Factories;

public class FrozenSetEnumerableFactoryTests
{
    [Fact]
    public void Begin_End_Success()
    {
        var factory = new FrozenSetEnumerableFactory<int>();

        // Begin.
        factory.Begin(1);
        var value = Assert.IsAssignableFrom<FrozenSet<int>>(factory.End());
        Assert.Equal([], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsAssignableFrom<FrozenSet<int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }

    [Fact]
    public void Begin_NegativeCount_ThrowsArgumentOutOfRangeException()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        Assert.Throws<ArgumentOutOfRangeException>("count", () => factory.Begin(-1));
    }

    [Fact]
    public void Add_End_Success()
    {
        var factory = new FrozenSetEnumerableFactory<int>();

        // Begin.
        factory.Begin(1);
        factory.Add(1);
        var value = Assert.IsAssignableFrom<FrozenSet<int>>(factory.End());
        Assert.Equal([1], value);

        // Begin again.
        factory.Begin(1);
        factory.Add(2);
        value = Assert.IsAssignableFrom<FrozenSet<int>>(factory.End());
        Assert.Equal([2], value);
    }

    [Fact]
    public void Add_OutOfRange_Success()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        factory.Begin(1);
        factory.Add(2);

        factory.Add(3);
        
        var value = Assert.IsAssignableFrom<FrozenSet<int>>(factory.End());
        Assert.Equal([2, 3], value);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        Assert.Throws<ExcelMappingException>(() => factory.Add(1));
    }

    [Fact]
    public void Set_Invoke_Success()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        factory.Begin(1);

        factory.Set(0, 1);
        Assert.Equal([1], Assert.IsAssignableFrom<FrozenSet<int>>(factory.End()));
    }

    [Fact]
    public void Set_InvokeOutOfRange_Success()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        factory.Begin(1);

        factory.Set(0, 1);
        factory.Set(5, 2);
        Assert.Equal([1, 0, 0, 0, 0, 2], Assert.IsAssignableFrom<FrozenSet<int>>(factory.End()));
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        Assert.Throws<ExcelMappingException>(() => factory.Set(0, 1));
    }

    [Fact]
    public void Set_NegativeIndex_ThrowsArgumentOutOfRangeException()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        factory.Begin(1);
        Assert.Throws<ArgumentOutOfRangeException>("index", () => factory.Set(-1, 1));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsAssignableFrom<FrozenSet<int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new FrozenSetEnumerableFactory<int>();
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsAssignableFrom<FrozenSet<int>>(factory.End());
        Assert.Equal([], value);
    }
}
