using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests;

public class EnumMapperTests
{
    [Theory]
    [InlineData(typeof(ConsoleColor))]
    public void Ctor_Type(Type enumType)
    {
        var mapper = new EnumMapper(enumType);
        Assert.Same(enumType, mapper.EnumType);
        Assert.False(mapper.IgnoreCase);
    }

    [Theory]
    [InlineData(typeof(ConsoleColor), true)]
    [InlineData(typeof(ConsoleColor), false)]
    public void Ctor_Type_Bool(Type enumType, bool ignoreCase)
    {
        var mapper = new EnumMapper(enumType, ignoreCase);
        Assert.Same(enumType, mapper.EnumType);
        Assert.Equal(ignoreCase, mapper.IgnoreCase);
    }

    [Fact]
    public void Ctor_NullEnumType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("enumType", () => new EnumMapper(null!));
    }

    [Theory]
    [InlineData(typeof(int))]
    [InlineData(typeof(Enum))]
    public void Ctor_EnumTypeNotEnum_ThrowsArgumentException(Type enumType)
    {
        Assert.Throws<ArgumentException>("enumType", () => new EnumMapper(enumType));
    }

    public static IEnumerable<object[]> Map_ValidStringValue_TestData()
    {
        yield return new object[] { new EnumMapper(typeof(ConsoleColor)), "Black", ConsoleColor.Black };
        yield return new object[] { new EnumMapper(typeof(ConsoleColor), ignoreCase: true), "Black", ConsoleColor.Black };
        yield return new object[] { new EnumMapper(typeof(ConsoleColor), ignoreCase: true), "bLaCk", ConsoleColor.Black };
    }

    [Theory]
    [MemberData(nameof(Map_ValidStringValue_TestData))]
    public void Map_ValidStringValue_ReturnsSuccess(EnumMapper mapper, string stringValue, Enum expected)
    {
        var result = mapper.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, result.Value);
        Assert.Null(result.Exception);
    }

    public static IEnumerable<object?[]> Map_InvalidStringValue_TestData()
    {
        yield return new object?[] { new EnumMapper(typeof(ConsoleColor)), null };
        yield return new object[] { new EnumMapper(typeof(ConsoleColor)), "" };
        yield return new object[] { new EnumMapper(typeof(ConsoleColor)), "Invalid" };
        yield return new object[] { new EnumMapper(typeof(ConsoleColor)), "black" };
    }

    [Theory]
    [MemberData(nameof(Map_InvalidStringValue_TestData))]
    public void Map_InvalidStringValue_ReturnsInvalid(EnumMapper mapper, string? stringValue)
    {
        var result = mapper.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.NotNull(result.Exception);
    }
}
