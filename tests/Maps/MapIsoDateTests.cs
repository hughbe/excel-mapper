using System.Globalization;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// ISO 8601 text is parsed with the one ISO format it can match before the culture's own patterns,
/// rather than trying each default format in turn. It must give exactly what trying the formats in
/// order gives, in every culture, so these tests compare the two.
/// </summary>
public class MapIsoDateTests
{
    // What the default formats accept for ISO text, as before the fast parser: the culture's "G"
    // pattern, then ISO 8601.
    private static readonly string[] s_isoFormats =
    [
        "yyyy-MM-dd",
        "yyyy-MM-ddTHH:mmK",
        "yyyy-MM-ddTHH:mm:ss.FFFFFFFK",
        "yyyy-MM-dd HH:mmK",
        "yyyy-MM-dd HH:mm:ss.FFFFFFFK",
    ];

    public static readonly string[] Texts =
    [
        "2024-03-04", "2024-02-29", "2023-02-29", "2024-13-01", "2024-00-10", "0001-01-01", "9999-12-31", "0000-01-01",
        "2024-03-04T10:30", "2024-03-04 10:30", "2024-03-04T10:30:15", "2024-03-04 10:30:15", "2024-03-04T23:59:59",
        "2024-03-04T24:00:00", "2024-03-04T10:60", "2024-03-04T10:30:60",
        "2024-03-04T10:30:15.1", "2024-03-04T10:30:15.1234567", "2024-03-04T10:30:15.12345678", "2024-03-04T10:30:15.",
        "2024-03-04T10:30Z", "2024-03-04T10:30:15Z", "2024-03-04T10:30:15.5Z", "2024-03-04T10:30:15+05:30",
        "2024-03-04T10:30:15-08:00", "2024-03-04 10:30:15+01:00", "2024-03-04T10:30:15+14:00", "2024-03-04T10:30:15-14:00",
        "2024-03-04T10:30:15+14:01", "2024-03-04T10:30:15+15:00", "2024-03-04T10:30:15+0530", "2024-03-04T10:30:15+05",
        "2024-03-04Z", "2024-03-04T10:30:15z", "2024-03-04t10:30", "2024-3-4", "24-03-04", "2024/03/04", "2024-03-04T1:30",
        "0001-01-01T00:00:00+01:00", "9999-12-31T23:59:59-01:00", "2024-03-31T01:30:00Z", "2024-10-27T01:30:00Z",
        "2024-03-10T02:30:00", "2024-11-03T01:30:00", " 2024-03-04", "2024-03-04 ", "2024-03-04T10:30:15 ",
    ];

    private static IEnumerable<CultureInfo> Cultures()
        => CultureInfo.GetCultures(CultureTypes.AllCultures).Where(c => !c.IsNeutralCulture || c == CultureInfo.InvariantCulture);

    private static ExcelImporter CreateImporter(CultureInfo culture)
    {
        var importer = Helpers.GetCsvImporter("Value\n" + string.Join("\n", Texts.Select(t => $"\"{t}\"")) + "\n");
        importer.Configuration.Culture = culture;
        return importer;
    }

    public class DateTimeRow
    {
        [ExcelOptional]
        public DateTime Value { get; set; }
    }

    public class DateTimeOffsetRow
    {
        public DateTimeOffset Value { get; set; }
    }

    [Fact]
    public void ReadRows_IsoDateTime_SameAsFormatsInEveryCulture()
    {
        foreach (var culture in Cultures())
        {
            using var importer = CreateImporter(culture);
            var results = importer.ReadSheet().ReadRowResults<DateTimeRow>().ToArray();
            for (int i = 0; i < Texts.Length; i++)
            {
                var text = Texts[i];
                var expectedValid = DateTime.TryParseExact(text, "G", culture, DateTimeStyles.AdjustToUniversal, out var expected)
                    || DateTime.TryParseExact(text, s_isoFormats, culture, DateTimeStyles.AdjustToUniversal, out expected);
                if (!expectedValid)
                {
                    // Other default formats, such as the culture's short date, may still read it.
                    continue;
                }

                Assert.True(results[i].IsValid, $"{text} in {culture.Name}");
                Assert.Equal((expected, expected.Kind), (results[i].Value!.Value, results[i].Value!.Value.Kind));
            }
        }
    }

    [Fact]
    public void ReadRows_IsoDateTimeOffset_SameAsFormatsInEveryCulture()
    {
        foreach (var culture in Cultures())
        {
            using var importer = CreateImporter(culture);
            var results = importer.ReadSheet().ReadRowResults<DateTimeOffsetRow>().ToArray();
            for (int i = 0; i < Texts.Length; i++)
            {
                var text = Texts[i];
                if (DateTimeOffset.TryParseExact(text, "G", culture, DateTimeStyles.AssumeUniversal, out var expected)
                    || DateTimeOffset.TryParseExact(text, s_isoFormats, culture, DateTimeStyles.AssumeUniversal, out expected))
                {
                    Assert.True(results[i].IsValid, $"{text} in {culture.Name}");
                    Assert.Equal((expected.DateTime, expected.Offset), (results[i].Value!.Value.DateTime, results[i].Value!.Value.Offset));
                }
            }
        }
    }

    [Fact]
    public void ReadRows_IsoTextInvalidEverywhere_IsInvalid()
    {
        // Text that looks like ISO 8601 but isn't a valid date or time is still invalid.
        foreach (var text in new[] { "2023-02-29", "2024-13-01", "2024-03-04T24:00:00", "2024-03-04T10:30:15+15:00" })
        {
            using var importer = Helpers.GetCsvImporter($"Value\n{text}\n");
            Assert.False(importer.ReadSheet().ReadRowResults<DateTimeRow>().Single().IsValid, text);
        }
    }

#if LIBRARY_NET8
    public class DateOnlyRow
    {
        public DateOnly Value { get; set; }
    }

    [Fact]
    public void ReadRows_IsoDateOnly_SameAsFormatsInEveryCulture()
    {
        foreach (var culture in Cultures())
        {
            using var importer = CreateImporter(culture);
            var results = importer.ReadSheet().ReadRowResults<DateOnlyRow>().ToArray();
            for (int i = 0; i < Texts.Length; i++)
            {
                var text = Texts[i];
                if (DateOnly.TryParseExact(text, "d", culture, DateTimeStyles.None, out var expected)
                    || DateOnly.TryParseExact(text, "yyyy-MM-dd", culture, DateTimeStyles.None, out expected))
                {
                    Assert.True(results[i].IsValid, $"{text} in {culture.Name}");
                    Assert.Equal(expected, results[i].Value!.Value);
                }
            }
        }
    }
#endif

    [Fact]
    public void ReadRows_IsoWithFormats_UsesFormatsOnly()
    {
        // Formats set on the member replace the defaults, ISO 8601 included.
        using var importer = Helpers.GetCsvImporter("Value\n2024-03-04\n");
        importer.Configuration.RegisterClassMap<DateTimeRow>(c => c.Map(p => p.Value).WithFormats("dd/MM/yyyy"));
        Assert.False(importer.ReadSheet().ReadRowResults<DateTimeRow>().Single().IsValid);
    }
}
