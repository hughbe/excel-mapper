using Xunit;

namespace ExcelMapper.Tests;

public class ExcelPreserveFormattingAttributeTests
{
    [Fact]
    public void Ctor_Default()
    {
        var exception = Record.Exception(() => new ExcelPreserveFormattingAttribute());
        Assert.Null(exception);
    }
}
