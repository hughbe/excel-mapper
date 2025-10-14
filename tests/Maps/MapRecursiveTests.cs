using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper;

public class MapRecursiveTests
{
    [Fact]
    public void ReadRows_RecursiveProperty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<RecursivePropertyClass>());
    }

    private class RecursivePropertyClass
    {
        public RecursivePropertyClass StringValue { get; set; } = default!;

        public string MappedValue { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_RecursiveField_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<RecursiveFieldClass>());
    }

    private class RecursiveFieldClass
    {
        public RecursiveFieldClass StringValue = default!;

        public string MappedValue = default!;
    }

    [Fact]
    public void ReadRows_RecursiveIndirect_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<RecursiveIndirectParent>());
    }

    private class RecursiveIndirectParent
    {
        public RecursiveIndirectChild B { get; set; } = default!;
    }

    private class RecursiveIndirectChild
    {
        public RecursiveIndirectParent A { get; set; } = default!;
    }
}