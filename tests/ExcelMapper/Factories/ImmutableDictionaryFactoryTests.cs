using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using Xunit;

namespace ExcelMapper.Factories;

public class ImmutableDictionaryFactoryTests
{
    [Fact]
    public void Begin_End_Success()
    {
        var factory = new ImmutableDictionaryFactory<int>();

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<ImmutableDictionary<string, int>>(factory.End());
        Assert.Equal([], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<ImmutableDictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new ImmutableDictionaryFactory<int>();
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }
    [Fact]
    public void Add_End_Success()
    {
        var factory = new ImmutableDictionaryFactory<int>();

        // Begin.
        factory.Begin(1);
        factory.Add("key", 1);
        var value = Assert.IsType<ImmutableDictionary<string, int>>(factory.End());
        Assert.Equal(ImmutableDictionary.CreateRange<string, int>(new Dictionary<string, int> { ["key"] = 1 }), value);

        // Begin again.
        factory.Begin(1);
        factory.Add("key", 2);
        value = Assert.IsType<ImmutableDictionary<string, int>>(factory.End());
        Assert.Equal(ImmutableDictionary.CreateRange<string, int>(new Dictionary<string, int> { ["key"] = 2 }), value);
    }

    [Fact]
    public void Add_OutOfRange_Success()
    {
        var factory = new ImmutableDictionaryFactory<int>();
        factory.Begin(1);
        factory.Add("key1", 2);

        factory.Add("key2", 3);
        
        var value = Assert.IsType<ImmutableDictionary<string, int>>(factory.End());
        Assert.Equal(ImmutableDictionary.CreateRange<string, int>(new Dictionary<string, int> { ["key1"] = 2, ["key2"] = 3 }), value);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ImmutableDictionaryFactory<int>();
        Assert.Throws<ExcelMappingException>(() => factory.Add("key", 1));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new ImmutableDictionaryFactory<int>();
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new ImmutableDictionaryFactory<int>();
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new ImmutableDictionaryFactory<int>();
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<ImmutableDictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new ImmutableDictionaryFactory<int>();
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<ImmutableDictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }
}
