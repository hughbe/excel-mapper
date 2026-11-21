using System.Linq;

namespace ExcelMapper.Tests;

public class MapCustomColumnLetterTests
{
    [Fact]
    public void ReadRows_AutoMappedColumnLetterProperty_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("x,y,z\n1,2,3\n");
        var row = importer.ReadSheet().ReadRows<ColumnLetterRow>().Single();
        Assert.Equal("3", row.Third);
        Assert.Equal("1", row.First);
    }

    public class ColumnLetterRow
    {
        [ExcelColumnLetter("C")]
        public string Third { get; set; } = "";

        [ExcelColumnLetter("a")]
        public string First { get; set; } = "";
    }

    [Fact]
    public void ReadRows_CustomMappedColumnLetter_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("x,y\n1,2\n");
        importer.Configuration.RegisterClassMap<FluentRow>(c => c.Map(p => p.Value).WithColumnLetter("b"));
        Assert.Equal("2", importer.ReadSheet().ReadRows<FluentRow>().Single().Value);
    }

    public class FluentRow
    {
        public string Value { get; set; } = "";
    }
}
