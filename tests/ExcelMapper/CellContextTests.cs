using System.Globalization;

namespace ExcelMapper.Tests;

public class CellContextTests
{
    [Fact]
    public void Ctor_Invoke_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var sheet = importer.ReadSheet();
        var member = typeof(CellContextTests).GetMethod(nameof(Ctor_Invoke_ReturnsExpected))!;
        var context = new CellContext(sheet, 10, 1, member);
        Assert.Same(sheet, context.Sheet);
        Assert.Equal(10, context.RowIndex);
        Assert.Equal(1, context.ColumnIndex);
        Assert.Same(member, context.Member);
        Assert.Equal("B11", context.CellAddress);
        Assert.Same(CultureInfo.CurrentCulture, context.FormatProvider);
    }

    [Fact]
    public void Ctor_NullMember_Success()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var context = new CellContext(importer.ReadSheet(), 0, 0, null);
        Assert.Null(context.Member);
        Assert.Equal("A1", context.CellAddress);
    }

    [Fact]
    public void Ctor_NullSheet_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("sheet", () => new CellContext(null!, 0, 0, null));
    }

    [Fact]
    public void Ctor_NegativeIndex_ThrowsArgumentOutOfRangeException()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var sheet = importer.ReadSheet();
        Assert.Throws<ArgumentOutOfRangeException>("rowIndex", () => new CellContext(sheet, -1, 0, null));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => new CellContext(sheet, 0, -1, null));
    }

    [Fact]
    public void FormatProvider_ImporterCulture_ReturnsCulture()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var culture = new CultureInfo("fr-FR");
        importer.Configuration.Culture = culture;
        Assert.Same(culture, new CellContext(importer.ReadSheet(), 0, 0, null).FormatProvider);
    }
}
