namespace ExcelMapper.Tests;

public class ExcelIgnoreAttributeTests
{
    [Fact]
    public void Ctor_Default()
    {
        var exception = Record.Exception(() => new ExcelIgnoreAttribute());
        Assert.Null(exception);
    }
}
