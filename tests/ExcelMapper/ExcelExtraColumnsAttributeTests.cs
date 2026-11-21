namespace ExcelMapper.Tests;

public class ExcelExtraColumnsAttributeTests
{
    [Fact]
    public void Ctor_Default()
    {
        var exception = Record.Exception(() => new ExcelExtraColumnsAttribute());
        Assert.Null(exception);
    }
}
