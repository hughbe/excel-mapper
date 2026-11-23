using System.Globalization;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// Dates and times written as text, as in CSV files. Without set formats, ISO 8601 and the culture's
/// short date, short date and time, and time patterns are accepted, as well as the default format.
/// </summary>
public class MapDateTextTests
{
    [Fact]
    public void ReadRows_IsoDates_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Date,NullableDate,Offset\n2024-03-04,2024-03-04T10:30:15,2024-03-04T10:30:15+01:00\n2024-12-31 23:59,,2024-12-31T00:00:00Z\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("en-GB");
        var rows = importer.ReadSheet().ReadRows<DateRow>().ToArray();
        Assert.Equal(new DateTime(2024, 3, 4), rows[0].Date);
        Assert.Equal(new DateTime(2024, 3, 4, 10, 30, 15), rows[0].NullableDate);
        Assert.Equal(new DateTimeOffset(2024, 3, 4, 10, 30, 15, TimeSpan.FromHours(1)), rows[0].Offset);
        Assert.Equal(TimeSpan.FromHours(1), rows[0].Offset.Offset);
        Assert.Equal(new DateTime(2024, 12, 31, 23, 59, 0), rows[1].Date);
        Assert.Null(rows[1].NullableDate);
        Assert.Equal(TimeSpan.Zero, rows[1].Offset.Offset);
    }

    [Theory]
    [InlineData("en-GB", "04/03/2024")]
    [InlineData("en-US", "3/4/2024")]
    [InlineData("de-DE", "04.03.2024")]
    [InlineData("ja-JP", "2024/03/04")]
    [InlineData("fr-FR", "2024-03-04")]
    public void ReadRows_CultureShortDate_ReturnsExpected(string culture, string date)
    {
        using var importer = Helpers.GetCsvImporter($"Date,NullableDate,Offset\n\"{date}\",\"{date}\",\"{date}\"\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo(culture);
        var row = importer.ReadSheet().ReadRows<DateRow>().Single();
        Assert.Equal(new DateTime(2024, 3, 4), row.Date);
        Assert.Equal(new DateTime(2024, 3, 4), row.NullableDate);
        Assert.Equal(new DateTime(2024, 3, 4), row.Offset.DateTime);
    }

    [Theory]
    [InlineData("3/4/2024 2:30 PM")]
    [InlineData("3/4/2024 2:30:00 PM")]
    public void ReadRows_UsDateAndTimeWithOrdinarySpace_ReturnsExpected(string dateTime)
    {
        // ICU's en-US patterns put U+202F before AM/PM; text with an ordinary space is accepted too.
        using var importer = Helpers.GetCsvImporter($"Date,NullableDate,Offset\n{dateTime},{dateTime},{dateTime}\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("en-US");
        var row = importer.ReadSheet().ReadRows<DateRow>().Single();
        Assert.Equal(new DateTime(2024, 3, 4, 14, 30, 0), row.Date);
        Assert.Equal(new DateTime(2024, 3, 4, 14, 30, 0), row.NullableDate);
    }

    public class DateRow
    {
        public DateTime Date { get; set; }
        public DateTime? NullableDate { get; set; }
        public DateTimeOffset Offset { get; set; }
    }

    [Theory]
    [InlineData("2024-3-4")]
    [InlineData("10:30")]
    [InlineData("1/2")]
    [InlineData("tomorrow")]
    public void ReadRows_InvalidDateText_ThrowsExcelMappingException(string value)
    {
        using var importer = Helpers.GetCsvImporter($"Date\n{value}\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("en-GB");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<SingleDateRow>().ToArray());
        Assert.Equal("A2", exception.CellAddress);
    }

    [Fact]
    public void ReadRows_ExcelFormatsAttribute_OnlyAcceptsThoseFormats()
    {
        using var importer = Helpers.GetCsvImporter("Date\n04/03/2024\n2024-03-04\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("en-GB");
        var results = importer.ReadSheet().ReadRowResults<FormatsRow>().ToArray();
        Assert.Equal(new DateTime(2024, 3, 4), results[0].Value!.Date);
        Assert.Equal("A3", Assert.Single(results[1].Errors).CellAddress);
    }

    [Fact]
    public void ReadRows_WithFormats_OnlyAcceptsThoseFormats()
    {
        using var importer = Helpers.GetCsvImporter("Date\n04/03/2024\n2024-03-04\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("en-GB");
        importer.Configuration.RegisterClassMap<SingleDateRow>(c => c.Map(r => r.Date).WithFormats("dd/MM/yyyy"));
        var results = importer.ReadSheet().ReadRowResults<SingleDateRow>().ToArray();
        Assert.Equal(new DateTime(2024, 3, 4), results[0].Value!.Date);
        Assert.False(results[1].IsValid);
    }

    [Fact]
    public void ReadRows_TypedDateCell_Unchanged()
    {
        var date = new DateTime(2024, 3, 4, 10, 30, 0);
        using var importer = XlsxBuilder.Create(["Date"], [date], ["2024-03-04"]);
        var rows = importer.ReadSheet().ReadRows<SingleDateRow>().ToArray();
        Assert.Equal(date, rows[0].Date);
        Assert.Equal(new DateTime(2024, 3, 4), rows[1].Date);
    }

    public class SingleDateRow
    {
        public DateTime Date { get; set; }
    }

    [Theory]
    [InlineData("2024-03-04T10:30:15Z", 10, DateTimeKind.Utc)]
    [InlineData("2024-03-04T10:30:15+01:00", 9, DateTimeKind.Utc)]
    [InlineData("2024-03-04T10:30:15-08:00", 18, DateTimeKind.Utc)]
    [InlineData("2024-03-04T10:30:15", 10, DateTimeKind.Unspecified)]
    [InlineData("03/04/2024 10:30:15", 10, DateTimeKind.Unspecified)]
    public void ReadRows_DateTimeTextWithOffset_ConvertsToUtc(string text, int hour, DateTimeKind kind)
    {
        // Not to the machine's local time, so every machine reads the same value.
        using var importer = Helpers.GetCsvImporter($"Date\n{text}\n");
        var date = importer.ReadSheet().ReadRows<SingleDateRow>().Single().Date;
        Assert.Equal(new DateTime(2024, 3, 4, hour, 30, 15), date);
        Assert.Equal(kind, date.Kind);
    }

    public class FormatsRow
    {
        [ExcelFormats("dd/MM/yyyy")]
        public DateTime Date { get; set; }
    }

#if LIBRARY_NET8
    [Fact]
    public void ReadRows_IsoDateOnlyAndTimeOnly_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Day,Time\n2024-03-04,14:30\n2024-12-31,9:05:30\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("en-US");
        var rows = importer.ReadSheet().ReadRows<DateOnlyRow>().ToArray();
        Assert.Equal(new DateOnly(2024, 3, 4), rows[0].Day);
        Assert.Equal(new TimeOnly(14, 30), rows[0].Time);
        Assert.Equal(new DateOnly(2024, 12, 31), rows[1].Day);
        Assert.Equal(new TimeOnly(9, 5, 30), rows[1].Time);
    }

    [Fact]
    public void ReadRows_CultureDateOnlyAndTimeOnly_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Day,Time\n3/4/2024,2:30 PM\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("en-US");
        var row = importer.ReadSheet().ReadRows<DateOnlyRow>().Single();
        Assert.Equal(new DateOnly(2024, 3, 4), row.Day);
        Assert.Equal(new TimeOnly(14, 30), row.Time);
    }

    public class DateOnlyRow
    {
        public DateOnly Day { get; set; }
        public TimeOnly Time { get; set; }
    }
#endif
}
