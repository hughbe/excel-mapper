namespace ExcelMapper.Tests;

public class ColumnSelectorContextTests
{
    [Fact]
    public void Ctor_Sheet_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("A,B\n1,2\n");
        var sheet = importer.ReadSheet();
        var context = new ColumnSelectorContext(sheet);
        Assert.Same(sheet, context.Sheet);
        Assert.Null(context.Heading);
        Assert.Equal(2, context.NumberOfColumns);

        var heading = sheet.ReadHeading();
        Assert.Same(heading, new ColumnSelectorContext(sheet).Heading);
    }

    [Fact]
    public void Ctor_NullSheet_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("sheet", () => new ColumnSelectorContext(null!));
    }
}
