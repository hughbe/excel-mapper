using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Tests;

public class IToOneMapExtensionsTests
{
    [Fact]
    public void MakeOptional_HasMapper_ReturnsExpected()
    {
        var map = new CustomToOneMap();
        Assert.False(map.Optional);
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
    }

    [Fact]
    public void MakeOptional_AlreadyOptional_ReturnsExpected()
    {
        var map = new CustomToOneMap();
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
    }

    [Fact]
    public void MakePreserveFormatting_HasMapper_ReturnsExpected()
    {
        var map = new CustomToOneMap();
        Assert.False(map.PreserveFormatting);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
    }

    [Fact]
    public void MakePreserveFormatting_AlreadyPreserveFormatting_ReturnsExpected()
    {
        var map = new CustomToOneMap();
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
    }

    private class CustomToOneMap : IToOneMap
    {
        public bool Optional { get; set; }
        public bool PreserveFormatting { get; set; }

        public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
            => throw new System.NotImplementedException();
    }
}
