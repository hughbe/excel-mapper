namespace ExcelMapper.Tests;

public class ExcelTrimStringAttributeTests
{
    [Fact]
    public void Ctor_Default()
    {
        var exception = Record.Exception(() => new ExcelTrimStringAttribute());
        Assert.Null(exception);
    }
}
