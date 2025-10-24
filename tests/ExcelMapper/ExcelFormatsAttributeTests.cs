namespace ExcelMapper.Tests;

public class ExcelFormatsAttributeTests
{
    public static IEnumerable<object?[]> Ctor_Formats_TestData()
    {
        yield return new object?[] { new string[] { "MM/dd/yyyy" } };
        yield return new object?[] { new string[] { "MM/dd/yyyy", "dd-MM-yyyy" } };
        yield return new object?[] { new string[] { "MM/dd/yyyy", "MM/dd/yyyy" } };
    }

    [Theory]
    [MemberData(nameof(Ctor_Formats_TestData))]
    public void Ctor_StringArray(string[] formats)
    {
        var attribute = new ExcelFormatsAttribute(formats);
        Assert.Same(formats, attribute.Formats);
    }

    [Fact]
    public void Ctor_NullFormats_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("formats", () => new ExcelFormatsAttribute(null!));
    }

    [Fact]
    public void Ctor_EmptyFormats_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("formats", () => new ExcelFormatsAttribute([]));
    }

    [Fact]
    public void Ctor_NullValueInFormats_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("formats", () => new ExcelFormatsAttribute(["MM/dd/yyyy", null!]));
    }

    [Fact]
    public void Ctor_EmptyStringInFormats_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("formats", () => new ExcelFormatsAttribute(["MM/dd/yyyy", ""]));
    }
}
