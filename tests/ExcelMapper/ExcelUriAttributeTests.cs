namespace ExcelMapper.Tests;

public class ExcelUriAttributeTests
{
    public static IEnumerable<object?[]> Ctor_UriKind_TestData()
    {
        yield return new object?[] { UriKind.Absolute };
        yield return new object?[] { UriKind.Relative };
        yield return new object?[] { UriKind.RelativeOrAbsolute };
    }

    [Theory]
    [MemberData(nameof(Ctor_UriKind_TestData))]
    public void Ctor_UriKind(UriKind uriKind)
    {
        var attribute = new ExcelUriAttribute(uriKind);
        Assert.Equal(uriKind, attribute.UriKind);
    }

    [Theory]
    [InlineData((UriKind)(-1))]
    [InlineData((UriKind)3)]
    [InlineData((UriKind)100)]
    public void Ctor_InvalidUriKind_ThrowsArgumentOutOfRangeException(UriKind uriKind)
    {
        Assert.Throws<ArgumentOutOfRangeException>("uriKind", () => new ExcelUriAttribute(uriKind));
    }
}
