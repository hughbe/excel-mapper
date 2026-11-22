namespace ExcelMapper.Tests;

public class ExcelColumnLetterAttributeTests
{
    [Theory]
    [InlineData("A", 0, "A")]
    [InlineData("a", 0, "A")]
    [InlineData("C", 2, "C")]
    [InlineData(" Z ", 25, "Z")]
    [InlineData("AA", 26, "AA")]
    [InlineData("ab", 27, "AB")]
    [InlineData("XFD", 16383, "XFD")]
    [InlineData("$B", 1, "B")]
    public void Ctor_String(string letter, int expectedIndex, string expectedLetter)
    {
        var attribute = new ExcelColumnLetterAttribute(letter);
        Assert.Equal(expectedIndex, attribute.Index);
        Assert.Equal(expectedLetter, attribute.Letter);
    }

    [Fact]
    public void Ctor_NullLetter_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("letter", () => new ExcelColumnLetterAttribute(null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("1")]
    [InlineData("A1")]
    [InlineData("$")]
    [InlineData("XFE")]
    [InlineData("AAAA")]
    [InlineData("É")]
    public void Ctor_InvalidLetter_ThrowsArgumentException(string letter)
    {
        Assert.Throws<ArgumentException>("letter", () => new ExcelColumnLetterAttribute(letter));
    }

    [Fact]
    public void ExcelColumnLetterAttribute_IsExcelColumnIndexAttribute()
    {
        Assert.IsAssignableFrom<ExcelColumnIndexAttribute>(new ExcelColumnLetterAttribute("A"));
    }
}
