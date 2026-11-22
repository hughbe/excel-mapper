namespace ExcelMapper.Tests;

public class ExcelColumnIndexAttributeTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(int.MaxValue)]
    public void Ctor_Int(int index)
    {
        var attribute = new ExcelColumnIndexAttribute(index);
        Assert.Equal(index, attribute.Index);
    }

    [Fact]
    public void Ctor_InvalidIndex_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("index", () => new ExcelColumnIndexAttribute(-1));
    }

}
