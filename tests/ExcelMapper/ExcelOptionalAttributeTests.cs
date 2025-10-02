using Xunit;

namespace ExcelMapper.Tests;

public class ExcelOptionalAttributeTests
{
    [Fact]
    public void Ctor_Default()
    {
        var exception = Record.Exception(() => new ExcelOptionalAttribute());
        Assert.Null(exception);
    }
}
