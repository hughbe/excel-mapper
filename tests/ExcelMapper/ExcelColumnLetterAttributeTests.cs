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

    [Theory]
    [InlineData(0, "A")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    public void Index_Set_GetReturnsExpected(int value, string expectedLetter)
    {
        var attribute = new ExcelColumnLetterAttribute("C")
        {
            Index = value
        };
        Assert.Equal(value, attribute.Index);
        Assert.Equal(expectedLetter, attribute.Letter);

        // Set same.
        attribute.Index = value;
        Assert.Equal(value, attribute.Index);
        Assert.Equal(expectedLetter, attribute.Letter);
    }

    [Fact]
    public void Index_SetInvalidValue_ThrowsArgumentOutOfRangeException()
    {
        var attribute = new ExcelColumnLetterAttribute("A");
        Assert.Throws<ArgumentOutOfRangeException>("value", () => attribute.Index = -1);
    }

    [Fact]
    public void ExcelColumnLetterAttribute_IsExcelColumnIndexAttribute()
    {
        Assert.IsAssignableFrom<ExcelColumnIndexAttribute>(new ExcelColumnLetterAttribute("A"));
    }
}
