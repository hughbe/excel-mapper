using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class ExcelColumnMatchingAttributeTests
{
    [Fact]
    public void Ctor_Type()
    {
        var matcherType = typeof(ColumnMatcher);
        var attribute = new ExcelColumnMatchingAttribute(matcherType);
        Assert.Same(matcherType, attribute.Type);
        Assert.Null(attribute.ConstructorArguments);
    }

    [Fact]
    public void Ctor_NullMatcherType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("matcherType", () => new ExcelColumnMatchingAttribute((Type)null!));
    }

    [Fact]
    public void Ctor_InvalidMatcherType_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("matcherType", () => new ExcelColumnMatchingAttribute(typeof(IExcelColumnMatcher)));
    }
    
    [Fact]
    public void Ctor_String_RegexOptions_Default()
    {
        var attribute = new ExcelColumnMatchingAttribute(@"Year \d+$");
        Assert.Equal(typeof(RegexColumnMatcher), attribute.Type);
        var regex = Assert.IsType<Regex>(Assert.Single(attribute.ConstructorArguments!));
        Assert.Matches(regex, "Year 2024");
        Assert.DoesNotMatch(regex, "year 2024");
    }
    
    [Fact]
    public void Ctor_String_RegexOptions_IgnoreCase()
    {
        var attribute = new ExcelColumnMatchingAttribute(@"Year \d+$", RegexOptions.IgnoreCase);
        Assert.Equal(typeof(RegexColumnMatcher), attribute.Type);
        var regex = Assert.IsType<Regex>(Assert.Single(attribute.ConstructorArguments!));
        Assert.Matches(regex, "Year 2024");
        Assert.Matches(regex, "year 2024");
    }
    
    [Fact]
    public void Ctor_NullPattern_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("pattern", () => new ExcelColumnMatchingAttribute((string)null!));
    }
    
    [Fact]
    public void Ctor_EmptyPattern_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("pattern", () => new ExcelColumnMatchingAttribute(string.Empty));
    }

    public static IEnumerable<object?[]> ConstructorArguments_Set_TestData()
    {
        yield return new object?[] { null };
        yield return new object[] { new object[0] };
        yield return new object[] { new object?[] { "Value", null } };
    }

    [Theory]
    [MemberData(nameof(ConstructorArguments_Set_TestData))]
    public void ConstructorArguments_Set_GetReturnsExpected(object?[]? value)
    {
        var attribute = new ExcelColumnMatchingAttribute(typeof(ColumnMatcher))
        {
            ConstructorArguments = value
        };
        Assert.Same(value, attribute.ConstructorArguments);
        
        // Set.
        attribute.ConstructorArguments = value;
        Assert.Same(value, attribute.ConstructorArguments);
    }

    private class ColumnMatcher : IExcelColumnMatcher
    {
        public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
            => throw new NotImplementedException();
    }
}
