using System.Globalization;

namespace ExcelMapper.Tests;

public class ExcelNumberStyleAttributeTests
{
    public static IEnumerable<object?[]> Ctor_NumberStyles_TestData()
    {
        yield return new object?[] { NumberStyles.None };
        yield return new object?[] { NumberStyles.Integer };
        yield return new object?[] { NumberStyles.Number };
        yield return new object?[] { NumberStyles.HexNumber };
        yield return new object?[] { NumberStyles.AllowThousands };
        yield return new object?[] { NumberStyles.AllowParentheses };
        yield return new object?[] { NumberStyles.AllowThousands | NumberStyles.AllowParentheses };
    }

    [Theory]
    [MemberData(nameof(Ctor_NumberStyles_TestData))]
    public void Ctor_NumberStyles(NumberStyles style)
    {
        var attribute = new ExcelNumberStyleAttribute(style);
        Assert.Equal(style, attribute.Style);
    }
}
