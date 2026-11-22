using System.Globalization;
using System.Linq;
using static ExcelMapper.Tests.XlsxBuilder;

namespace ExcelMapper.Tests;

public class CultureTests
{
    public class Row
    {
        public double Amount { get; set; }
        public DateTime When { get; set; }
        public string Text { get; set; } = "";
    }

    private static readonly CultureInfo s_german = new("de-DE");

    [Fact]
    public void Culture_Default_IsInvariant()
    {
        using var importer = Helpers.GetCsvImporter("a\n1\n");
        Assert.Same(CultureInfo.InvariantCulture, importer.Configuration.Culture);
    }

    [Fact]
    public void ReadRows_CultureSet_ParsesTextWithCulture()
    {
        var date = new DateTime(2024, 3, 4, 5, 6, 7);
        using var importer = Helpers.GetCsvImporter($"Amount;When;Text\n\"1,5\";{date.ToString("G", s_german)};x\n");
        importer.Configuration.Culture = s_german;

        var row = WithCulture(CultureInfo.InvariantCulture, () => importer.ReadSheet().ReadRows<Row>().Single());
        Assert.Equal(1.5, row.Amount);
        Assert.Equal(date, row.When);
    }

    [Fact]
    public void ReadRows_CultureSet_ConvertsNumbersToTextWithCulture()
    {
        using var importer = Create(["Amount", "When", "Text"], [1.5, new DateTime(2024, 3, 4), 2.5]);
        importer.Configuration.Culture = s_german;

        var row = WithCulture(CultureInfo.InvariantCulture, () => importer.ReadSheet().ReadRows<Row>().Single());
        Assert.Equal("2,5", row.Text);
    }

    [Fact]
    public void ReadRows_CultureNotSet_UsesInvariantCulture()
    {
        // The machine's culture doesn't change how text is read: "1.5" isn't fifteen in Germany.
        var date = new DateTime(2024, 3, 4, 5, 6, 7);
        using var importer = Helpers.GetCsvImporter($"Amount,When,Text\n1.5,{date.ToString(CultureInfo.InvariantCulture)},x\n");

        var row = WithCulture(s_german, () => importer.ReadSheet().ReadRows<Row>().Single());
        Assert.Equal(1.5, row.Amount);
        Assert.Equal(date, row.When);
    }

    [Fact]
    public void ReadRows_CultureNotSet_ConvertsNumbersToTextWithInvariantCulture()
    {
        using var importer = Create(["Amount", "When", "Text"], [1.5, new DateTime(2024, 3, 4), 2.5]);

        var row = WithCulture(s_german, () => importer.ReadSheet().ReadRows<Row>().Single());
        Assert.Equal("2.5", row.Text);
    }

    [Fact]
    public void ReadRows_CultureSetToCurrentCulture_UsesCurrentCulture()
    {
        using var importer = Helpers.GetCsvImporter("Amount;When;Text\n\"1,5\";;x\n");
        importer.Configuration.RegisterClassMap<Row>(c => c.Map(p => p.Amount));

        var row = WithCulture(s_german, () =>
        {
            importer.Configuration.Culture = CultureInfo.CurrentCulture;
            return importer.ReadSheet().ReadRows<Row>().Single();
        });
        Assert.Equal(1.5, row.Amount);
    }

    [Fact]
    public void ReadRows_CultureSet_LeavesTheThreadsCulture()
    {
        // The importer's culture is given to each cell, not set as the thread's culture: code
        // that runs while a row is mapped, such as a converter, sees the caller's culture, and
        // gets the importer's from its context.
        using var importer = Helpers.GetCsvImporter("Amount,Label\n1,a\nbad,b\n");
        importer.Configuration.Culture = s_german;
        var cultures = new List<(CultureInfo Thread, IFormatProvider Context)>();
        importer.Configuration.RegisterClassMap<LabelledRow>(c =>
        {
            c.Map(p => p.Amount);
            c.Map(p => p.Label).WithConverter(new RecordingConverter(cultures));
        });

        WithCulture(CultureInfo.InvariantCulture, () =>
        {
            var sheet = importer.ReadSheet();
            foreach (var row in sheet.ReadRows<LabelledRow>().Take(1))
            {
                Assert.Same(CultureInfo.InvariantCulture, CultureInfo.CurrentCulture);
            }

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<LabelledRow>());
            Assert.Same(CultureInfo.InvariantCulture, CultureInfo.CurrentCulture);
            return 0;
        });

        var (thread, context) = Assert.Single(cultures);
        Assert.Same(CultureInfo.InvariantCulture, thread);
        Assert.Same(s_german, context);
    }

    public class LabelledRow
    {
        public double Amount { get; set; }
        public string Label { get; set; } = default!;
    }

    private sealed class RecordingConverter(List<(CultureInfo Thread, IFormatProvider Context)> cultures) : ICellConverter<string>
    {
        public CellConversionResult<string> Convert(CellValue value, CellContext context)
        {
            cultures.Add((CultureInfo.CurrentCulture, context.FormatProvider));
            return CellConversionResult<string>.Success(value.Text!);
        }
    }

#if LIBRARY_NET8
    [Fact]
    public void ReadRows_MemberFormatProvider_TakesPrecedenceOverCulture()
    {
        using var importer = Helpers.GetCsvImporter("Amount\n1.5\n");
        importer.Configuration.Culture = s_german;
        importer.Configuration.RegisterClassMap<Row>(c => c.Map(p => p.Amount).WithFormatProvider(CultureInfo.InvariantCulture));

        Assert.Equal(1.5, importer.ReadSheet().ReadRows<Row>().Single().Amount);
    }
#endif

    private static T WithCulture<T>(CultureInfo culture, Func<T> action)
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = culture;
            return action();
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }
}
