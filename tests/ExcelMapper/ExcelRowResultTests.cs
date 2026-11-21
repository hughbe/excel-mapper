using System.Linq;

namespace ExcelMapper.Tests;

public class ExcelRowResultTests
{
    [Fact]
    public void Default_ReturnsExpected()
    {
        ExcelRowResult<int> result = default;
        Assert.Equal(0, result.RowIndex);
        Assert.Equal(0, result.Value);
        Assert.Empty(result.Errors);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void Default_ReferenceType_ReturnsExpected()
    {
        ExcelRowResult<string> result = default;
        Assert.Null(result.Value);
        Assert.Empty(result.Errors);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void ValidRow_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        var result = importer.ReadSheet().ReadRowResults<int>().Single();

        Assert.Equal(1, result.RowIndex);
        Assert.Equal(1, result.Value);
        Assert.Empty(result.Errors);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void InvalidRow_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Value\ninvalid\n");
        var result = importer.ReadSheet().ReadRowResults<int>().Single();

        Assert.Equal(1, result.RowIndex);
        Assert.Equal(0, result.Value);
        var error = Assert.Single(result.Errors);
        Assert.Equal("A2", error.CellAddress);
        Assert.False(result.IsValid);
    }

    [Fact]
    public void ValidRows_ShareEmptyErrors()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n2\n");
        var results = importer.ReadSheet().ReadRowResults<int>().ToArray();
        Assert.Same(results[0].Errors, results[1].Errors);
    }
}
