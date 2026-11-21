namespace ExcelMapper.Tests;

public class ExcelRowNumberAttributeTests
{
    [Fact]
    public void Ctor_Default()
    {
        var exception = Record.Exception(() => new ExcelRowNumberAttribute());
        Assert.Null(exception);
    }
}
