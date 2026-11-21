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
    public void Culture_Default_IsNull()
    {
        using var importer = Helpers.GetCsvImporter("a\n1\n");
        Assert.Null(importer.Configuration.Culture);
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
    public void ReadRows_CultureNotSet_UsesCurrentCulture()
    {
        using var importer = Helpers.GetCsvImporter("Amount;When;Text\n\"1,5\";;x\n");
        importer.Configuration.RegisterClassMap<Row>(c => c.Map(p => p.Amount));

        var row = WithCulture(s_german, () => importer.ReadSheet().ReadRows<Row>().Single());
        Assert.Equal(1.5, row.Amount);
    }

    [Fact]
    public void ReadRows_CultureSet_RestoresCultureBetweenRowsAndAfterErrors()
    {
        using var importer = Helpers.GetCsvImporter("Amount\n1\nbad\n");
        importer.Configuration.Culture = s_german;
        importer.Configuration.RegisterClassMap<Row>(c => c.Map(p => p.Amount));

        WithCulture(CultureInfo.InvariantCulture, () =>
        {
            var sheet = importer.ReadSheet();
            foreach (var row in sheet.ReadRows<Row>().Take(1))
            {
                // The caller's code between rows runs in the caller's culture.
                Assert.Same(CultureInfo.InvariantCulture, CultureInfo.CurrentCulture);
            }

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Row>());
            Assert.Same(CultureInfo.InvariantCulture, CultureInfo.CurrentCulture);
            return 0;
        });
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
