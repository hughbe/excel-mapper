using System.Text.RegularExpressions;

namespace ExcelMapper.Tests;

public class ExcelColumnsMatchingAttributeTests
{
    [Theory]
    [InlineData(typeof(ColumnMatcher))]
    [InlineData(typeof(Selector))]
    [InlineData(typeof(NoConstructorExcelColumnMatcher))]
    public void Ctor_Type(Type matcherType)
    {
        var attribute = new ExcelColumnsMatchingAttribute(matcherType);
        Assert.Same(matcherType, attribute.Type);
        Assert.Null(attribute.ConstructorArguments);
    }

    [Fact]
    public void Ctor_NullMatcherType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("matcherType", () => new ExcelColumnsMatchingAttribute((Type)null!));
    }

    [Theory]
    [InlineData(typeof(IColumnSelector))]
    [InlineData(typeof(ISubExcelColumnMatcher))]
    [InlineData(typeof(AbstractExcelColumnMatcher))]
    [InlineData(typeof(int))]
    [InlineData(typeof(object))]
    [InlineData(typeof(ExcelColumnsMatchingAttributeTests))]
    public void Ctor_InvalidMatcherType_ThrowsArgumentException(Type matcherType)
    {
        Assert.Throws<ArgumentException>("matcherType", () => new ExcelColumnsMatchingAttribute(matcherType));
    }

    private interface ISubExcelColumnMatcher : IColumnSelector
    {
    }

    private abstract class AbstractExcelColumnMatcher : IColumnSelector
    {
        public abstract IReadOnlyList<int> SelectColumns(ColumnSelectorContext context);

        public abstract string Describe(ColumnSelectorContext context);
    }

    private class NoConstructorExcelColumnMatcher : IColumnSelector
    {
        private NoConstructorExcelColumnMatcher()
        {
        }

        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context)
            => throw new NotImplementedException();

        public string Describe(ColumnSelectorContext context)
            => throw new NotImplementedException();
    }
    
    [Fact]
    public void Ctor_String_RegexOptions_Default()
    {
        var attribute = new ExcelColumnsMatchingAttribute(@"Year \d+$");
        Assert.Equal("RegexColumnMatcher", attribute.Type.Name);
        var regex = Assert.IsType<Regex>(Assert.Single(attribute.ConstructorArguments!));
        Assert.Equal(RegexOptions.None, regex.Options);
        Assert.Matches(regex, "Year 2024");
        Assert.DoesNotMatch(regex, "year 2024");
    }
    
    [Fact]
    public void Ctor_String_RegexOptions_IgnoreCase()
    {
        var attribute = new ExcelColumnsMatchingAttribute(@"Year \d+$", RegexOptions.IgnoreCase);
        Assert.Equal("RegexColumnMatcher", attribute.Type.Name);
        var regex = Assert.IsType<Regex>(Assert.Single(attribute.ConstructorArguments!));
        Assert.Equal(RegexOptions.IgnoreCase, regex.Options);
        Assert.Matches(regex, "Year 2024");
        Assert.Matches(regex, "year 2024");
    }
    
    [Fact]
    public void Ctor_NullPattern_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("pattern", () => new ExcelColumnsMatchingAttribute((string)null!));
    }
    
    [Fact]
    public void Ctor_EmptyPattern_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("pattern", () => new ExcelColumnsMatchingAttribute(string.Empty));
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
        var attribute = new ExcelColumnsMatchingAttribute(typeof(ColumnMatcher))
        {
            ConstructorArguments = value
        };
        Assert.Same(value, attribute.ConstructorArguments);
        
        // Set.
        attribute.ConstructorArguments = value;
        Assert.Same(value, attribute.ConstructorArguments);
    }

    private class ColumnMatcher : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context)
            => throw new NotImplementedException();

        public string Describe(ColumnSelectorContext context)
            => throw new NotImplementedException();
    }

    private sealed class Selector : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context) => [0];

        public string Describe(ColumnSelectorContext context) => "the first column";
    }
}
