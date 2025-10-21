using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Transformers.Tests;

public class TrimCellTransformerTests
{
    [Theory]
    [InlineData(null, null)]
    [InlineData("", "")]
    [InlineData(" abc ", "abc")]
    public void TransformStringValue_Invoke_ReturnsExpected(string? stringValue, string? expected)
    {
        var transformer = new TrimCellTransformer();
        Assert.Equal(expected, transformer.TransformStringValue(null!, 0, new ReadCellResult(0, stringValue, preserveFormatting: false)));
    }
}
