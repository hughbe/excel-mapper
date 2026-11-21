namespace ExcelMapper.Tests;

public class ExcelOffsetAttributeTests
{
    [Theory]
    [InlineData("+00:00", 0)]
    [InlineData("-00:00", 0)]
    [InlineData("00:00", 0)]
    [InlineData("+05:30", 330)]
    [InlineData("05:30", 330)]
    [InlineData("-08:00", -480)]
    [InlineData(" +14:00 ", 840)]
    [InlineData("-14:00", -840)]
    public void Ctor_String(string offset, int expectedMinutes)
    {
        var attribute = new ExcelOffsetAttribute(offset);
        Assert.Equal(TimeSpan.FromMinutes(expectedMinutes), attribute.Offset);
    }

    [Fact]
    public void Ctor_NullOffset_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("offset", () => new ExcelOffsetAttribute(null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("Z")]
    [InlineData("UTC")]
    [InlineData("5")]
    [InlineData("+5:30")]
    [InlineData("+05:60")]
    [InlineData("+15:00")]
    [InlineData("-14:01")]
    [InlineData("+05:30:00")]
    [InlineData("++05:30")]
    public void Ctor_InvalidOffset_ThrowsArgumentException(string offset)
    {
        Assert.Throws<ArgumentException>("offset", () => new ExcelOffsetAttribute(offset));
    }
}
