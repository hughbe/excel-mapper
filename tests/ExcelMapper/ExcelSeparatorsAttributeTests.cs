namespace ExcelMapper.Tests;

public class ExcelSeparatorsAttributeTests
{
    public static IEnumerable<object?[]> Ctor_StringSeparators_TestData()
    {
        yield return new object?[] { new string[] { "," } };
        yield return new object?[] { new string[] { ",", "|" } };
        yield return new object?[] { new string[] { ",", "," } };
    }

    [Theory]
    [MemberData(nameof(Ctor_StringSeparators_TestData))]
    public void Ctor_StringArray(string[] separators)
    {
        var attribute = new ExcelSeparatorsAttribute(separators);
        Assert.Null(attribute.CharSeparators);
        Assert.Same(separators, attribute.StringSeparators);
    }

    [Fact]
    public void Ctor_NullStringSeparators_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("separators", () => new ExcelSeparatorsAttribute((string[])null!));
    }

    [Fact]
    public void Ctor_EmptyStringSeparators_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("separators", () => new ExcelSeparatorsAttribute((string[])[]));
    }

    [Fact]
    public void Ctor_NullValueInStringSeparators_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("separators", () => new ExcelSeparatorsAttribute(["MM/dd/yyyy", null!]));
    }

    [Fact]
    public void Ctor_EmptyStringInStringSeparators_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("separators", () => new ExcelSeparatorsAttribute(["MM/dd/yyyy", ""]));
    }
    public static IEnumerable<object?[]> Ctor_CharSeparators_TestData()
    {
        yield return new object?[] { new char[] { ',' } };
        yield return new object?[] { new char[] { ',', '|' } };
        yield return new object?[] { new char[] { ',', ',' } };
    }

    [Theory]
    [MemberData(nameof(Ctor_CharSeparators_TestData))]
    public void Ctor_CharArray(char[] separators)
    {
        var attribute = new ExcelSeparatorsAttribute(separators);
        Assert.Same(separators, attribute.CharSeparators);
        Assert.Null(attribute.StringSeparators);
    }

    [Fact]
    public void Ctor_NullCharSeparators_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("separators", () => new ExcelSeparatorsAttribute((char[])null!));
    }

    [Fact]
    public void Ctor_EmptyCharSeparators_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("separators", () => new ExcelSeparatorsAttribute((char[])[]));
    }

    [Theory]
    [InlineData(StringSplitOptions.None - 1)]
    [InlineData(StringSplitOptions.None)]
    [InlineData(StringSplitOptions.RemoveEmptyEntries)]
    [InlineData(StringSplitOptions.RemoveEmptyEntries + 1)]
    public void Options_Set_GetReturnsExpected(StringSplitOptions value)
    {
        var attribute = new ExcelSeparatorsAttribute(',')
        {
            Options = value
        };
        Assert.Equal(value, attribute.Options);

        // Set same.
        attribute.Options = value;
        Assert.Equal(value, attribute.Options);
    }
}
