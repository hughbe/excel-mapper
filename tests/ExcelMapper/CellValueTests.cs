using System.Globalization;
using System.Linq;

namespace ExcelMapper.Tests;

public class CellValueTests
{
    [Theory]
    [InlineData("abc", CellKind.Text)]
    [InlineData("", CellKind.Empty)]
    [InlineData(null, CellKind.Empty)]
    public void Ctor_String(string? text, CellKind expectedKind)
    {
        var value = new CellValue(text);
        Assert.Equal(expectedKind, value.Kind);
        Assert.Equal(expectedKind == CellKind.Empty, value.IsEmpty);
        Assert.Equal(text, value.Text);
        Assert.Equal(text, value.Value);
        Assert.Equal(text ?? string.Empty, value.ToString());
    }

    [Fact]
    public void Default_Get_IsEmpty()
    {
        var value = default(CellValue);
        Assert.Equal(CellKind.Empty, value.Kind);
        Assert.True(value.IsEmpty);
        Assert.Null(value.Text);
        Assert.Null(value.Value);
    }

    [Fact]
    public void Number_Text_ThrowsInvalidOperationException()
    {
        var value = new CellValue("1");
        Assert.Throws<InvalidOperationException>(() => value.Number);
        Assert.Throws<InvalidOperationException>(() => value.Boolean);
        Assert.Throws<InvalidOperationException>(() => value.DateTime);
        Assert.Throws<InvalidOperationException>(() => value.TimeSpan);
    }

    [Fact]
    public void Kind_TypedCells_ReturnsExpected()
    {
        var date = new DateTime(2024, 3, 4, 10, 30, 0);
        using var importer = XlsxBuilder.Create(
            ["Number", "Boolean", "Date", "Duration", "Text", "Empty"],
            [2.5, true, date, new XlsxBuilder.Duration(TimeSpan.FromHours(30)), "abc", null]);
        importer.Configuration.RegisterClassMap<Captured>(c =>
        {
            c.Map(r => r.Number).WithConverter(new CapturingConverter());
            c.Map(r => r.Boolean).WithConverter(new CapturingConverter());
            c.Map(r => r.Date).WithConverter(new CapturingConverter());
            c.Map(r => r.Duration).WithConverter(new CapturingConverter());
            c.Map(r => r.Text).WithConverter(new CapturingConverter());
            c.Map(r => r.Empty).WithConverter(new CapturingConverter()).WithEmptyFallback(null);
        });
        var row = importer.ReadSheet().ReadRows<Captured>().Single();

        Assert.Equal(CellKind.Number, row.Number!.Value.Kind);
        Assert.Equal(2.5, row.Number.Value.Number);
        Assert.Equal(2.5, row.Number.Value.Value);
        Assert.Equal(2.5.ToString(CultureInfo.CurrentCulture), row.Number.Value.Text);

        Assert.Equal(CellKind.Boolean, row.Boolean!.Value.Kind);
        Assert.True(row.Boolean.Value.Boolean);

        Assert.Equal(CellKind.DateTime, row.Date!.Value.Kind);
        Assert.Equal(date, row.Date.Value.DateTime);

        Assert.Equal(CellKind.TimeSpan, row.Duration!.Value.Kind);
        Assert.Equal(TimeSpan.FromHours(30), row.Duration.Value.TimeSpan);

        Assert.Equal(CellKind.Text, row.Text!.Value.Kind);
        Assert.Equal("abc", row.Text.Value.Text);
        Assert.Throws<InvalidOperationException>(() => row.Text.Value.Number);

        // Empty cells use the empty fallback rather than the converter.
        Assert.Null(row.Empty);
    }

    [Fact]
    public void Kind_PreserveFormatting_IsText()
    {
        using var importer = XlsxBuilder.Create(["Number"], [2.5]);
        importer.Configuration.RegisterClassMap<Captured>(c => c.Map(r => r.Number).WithConverter(new CapturingConverter()).MakePreserveFormatting());
        var row = importer.ReadSheet().ReadRows<Captured>().Single();
        Assert.Equal(CellKind.Text, row.Number!.Value.Kind);
        Assert.Equal(row.Number.Value.Text, row.Number.Value.Value);
        Assert.Throws<InvalidOperationException>(() => row.Number.Value.Number);
    }

    public class Captured
    {
        public CellValue? Number { get; set; }
        public CellValue? Boolean { get; set; }
        public CellValue? Date { get; set; }
        public CellValue? Duration { get; set; }
        public CellValue? Text { get; set; }
        public CellValue? Empty { get; set; }
    }

    private sealed class CapturingConverter : ICellConverter<CellValue?>
    {
        public CellConversionResult<CellValue?> Convert(CellValue value, CellContext context) => CellConversionResult<CellValue?>.Success(value);
    }
}
