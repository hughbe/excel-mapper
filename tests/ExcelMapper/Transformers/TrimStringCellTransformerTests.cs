using ExcelMapper.Abstractions;

namespace ExcelMapper.Transformers.Tests;

public class TrimStringCellTransformerTests
{
    [Theory]
    [InlineData(null, null)]
    [InlineData("", "")]
    [InlineData(" abc ", "abc")]
    public void TransformStringValue_Invoke_ReturnsExpected(string? stringValue, string? expected)
    {
        var transformer = new TrimStringCellTransformer();
        Assert.Equal(expected, transformer.TransformStringValue(null!, 0, new ReadCellResult(0, stringValue, preserveFormatting: false)));
    }
}
