using System.Globalization;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// WithFormats and WithFormatProvider for dates, times and durations, and their nullable forms.
/// </summary>
public class MapDateFormatOptionsTests
{
    [Fact]
    public void ReadRows_NullableWithFormats_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("DateTime,Offset,TimeSpan\n20240304,20240304,01-30\n,,\n");
        importer.Configuration.RegisterClassMap<NullableRow>(c =>
        {
            c.Map(r => r.DateTime).WithFormats("yyyyMMdd");
            c.Map(r => r.Offset).WithFormats("yyyyMMdd");
            c.Map(r => r.TimeSpan).WithFormats(@"hh\-mm");
        });
        var rows = importer.ReadSheet().ReadRows<NullableRow>().ToArray();
        Assert.Equal(new DateTime(2024, 3, 4), rows[0].DateTime);
        Assert.Equal(new DateTime(2024, 3, 4), rows[0].Offset!.Value.DateTime);
        Assert.Equal(new TimeSpan(1, 30, 0), rows[0].TimeSpan);
        Assert.Null(rows[1].DateTime);
        Assert.Null(rows[1].Offset);
        Assert.Null(rows[1].TimeSpan);
    }

    [Fact]
    public void ReadRows_NullableWithFormatsEnumerable_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("DateTime,Offset,TimeSpan\n20240304,20240304,01-30\n");
        importer.Configuration.RegisterClassMap<NullableRow>(c =>
        {
            c.Map(r => r.DateTime).WithFormats(new List<string> { "yyyyMMdd" });
            c.Map(r => r.Offset).WithFormats(new List<string> { "yyyyMMdd" });
            c.Map(r => r.TimeSpan).WithFormats(new List<string> { @"hh\-mm" });
        });
        var row = importer.ReadSheet().ReadRows<NullableRow>().Single();
        Assert.Equal(new DateTime(2024, 3, 4), row.DateTime);
        Assert.Equal(new DateTime(2024, 3, 4), row.Offset!.Value.DateTime);
        Assert.Equal(new TimeSpan(1, 30, 0), row.TimeSpan);
    }

    [Fact]
    public void ReadRows_WithFormatProvider_UsesProvider()
    {
        // With the importer's en-US culture, 04/03/2024 would be 3 April, and 1:02:03,5 isn't a valid
        // duration in the culture-sensitive "g" format.
        using var importer = Helpers.GetCsvImporter("DateTime,Offset,TimeSpan\n04/03/2024,04/03/2024,\"1:02:03,5\"\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("en-US");
        var british = CultureInfo.GetCultureInfo("en-GB");
        var french = CultureInfo.GetCultureInfo("fr-FR");
        importer.Configuration.RegisterClassMap<Row>(c =>
        {
            c.Map(r => r.DateTime).WithFormatProvider(british);
            c.Map(r => r.Offset).WithFormatProvider(british);
            c.Map(r => r.TimeSpan).WithFormats("g").WithFormatProvider(french);
        });
        var row = importer.ReadSheet().ReadRows<Row>().Single();
        Assert.Equal(new DateTime(2024, 3, 4), row.DateTime);
        Assert.Equal(new DateTime(2024, 3, 4), row.Offset.DateTime);
        Assert.Equal(new TimeSpan(0, 1, 2, 3, 500), row.TimeSpan);
    }

    [Fact]
    public void ReadRows_NullableWithFormatProvider_UsesProvider()
    {
        using var importer = Helpers.GetCsvImporter("DateTime,Offset,TimeSpan\n04/03/2024,04/03/2024,\"1:02:03,5\"\n,,\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("en-US");
        var british = CultureInfo.GetCultureInfo("en-GB");
        var french = CultureInfo.GetCultureInfo("fr-FR");
        importer.Configuration.RegisterClassMap<NullableRow>(c =>
        {
            c.Map(r => r.DateTime).WithFormatProvider(british);
            c.Map(r => r.Offset).WithFormatProvider(british);
            c.Map(r => r.TimeSpan).WithFormats("g").WithFormatProvider(french);
        });
        var rows = importer.ReadSheet().ReadRows<NullableRow>().ToArray();
        Assert.Equal(new DateTime(2024, 3, 4), rows[0].DateTime);
        Assert.Equal(new DateTime(2024, 3, 4), rows[0].Offset!.Value.DateTime);
        Assert.Equal(new TimeSpan(0, 1, 2, 3, 500), rows[0].TimeSpan);
        Assert.Null(rows[1].DateTime);
        Assert.Null(rows[1].Offset);
        Assert.Null(rows[1].TimeSpan);
    }

    [Fact]
    public void ReadRows_ElementsWithFormatProvider_UsesProvider()
    {
        using var importer = Helpers.GetCsvImporter("Dates\n\"04/03/2024,05/03/2024\"\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("en-US");
        importer.Configuration.RegisterClassMap<ElementsRow>(c =>
            c.Map(r => r.Dates).WithElementMap(e => e.WithFormatProvider(CultureInfo.GetCultureInfo("en-GB"))));
        Assert.Equal([new DateTime(2024, 3, 4), new DateTime(2024, 3, 5)], importer.ReadSheet().ReadRows<ElementsRow>().Single().Dates);
    }

#if LIBRARY_NET8
    [Fact]
    public void ReadRows_DateOnlyTimeOnlyWithFormatProvider_UsesProvider()
    {
        // With the importer's de-DE culture, 04/03/2024 and 2:30 PM aren't valid.
        using var importer = Helpers.GetCsvImporter("Date,Time,NullableDate,NullableTime\n04/03/2024,2:30 PM,04/03/2024,2:30 PM\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("de-DE");
        var british = CultureInfo.GetCultureInfo("en-GB");
        var american = CultureInfo.GetCultureInfo("en-US");
        importer.Configuration.RegisterClassMap<DateOnlyRow>(c =>
        {
            c.Map(r => r.Date).WithFormatProvider(british);
            c.Map(r => r.Time).WithFormatProvider(american);
            c.Map(r => r.NullableDate).WithFormatProvider(british);
            c.Map(r => r.NullableTime).WithFormatProvider(american);
        });
        var row = importer.ReadSheet().ReadRows<DateOnlyRow>().Single();
        Assert.Equal(new DateOnly(2024, 3, 4), row.Date);
        Assert.Equal(new TimeOnly(14, 30), row.Time);
        Assert.Equal(new DateOnly(2024, 3, 4), row.NullableDate);
        Assert.Equal(new TimeOnly(14, 30), row.NullableTime);
    }

    [Fact]
    public void ReadRows_NullableDateOnlyTimeOnlyWithFormats_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Date,Time,NullableDate,NullableTime\n20240304,14-30,20240304,14-30\n20240305,15-45,,\n");
        importer.Configuration.RegisterClassMap<DateOnlyRow>(c =>
        {
            c.Map(r => r.Date).WithFormats("yyyyMMdd");
            c.Map(r => r.Time).WithFormats("HH-mm");
            c.Map(r => r.NullableDate).WithFormats("yyyyMMdd");
            c.Map(r => r.NullableTime).WithFormats(new List<string> { "HH-mm" });
        });
        var rows = importer.ReadSheet().ReadRows<DateOnlyRow>().ToArray();
        Assert.Equal(new DateOnly(2024, 3, 4), rows[0].NullableDate);
        Assert.Equal(new TimeOnly(14, 30), rows[0].NullableTime);
        Assert.Null(rows[1].NullableDate);
        Assert.Null(rows[1].NullableTime);

        using var enumerableImporter = Helpers.GetCsvImporter("Date,Time,NullableDate,NullableTime\n20240304,14-30,20240304,14-30\n");
        enumerableImporter.Configuration.RegisterClassMap<DateOnlyRow>(c =>
        {
            c.Map(r => r.Date).WithFormats("yyyyMMdd");
            c.Map(r => r.Time).WithFormats("HH-mm");
            c.Map(r => r.NullableDate).WithFormats(new List<string> { "yyyyMMdd" });
            c.Map(r => r.NullableTime).WithFormats("HH-mm");
        });
        var row = enumerableImporter.ReadSheet().ReadRows<DateOnlyRow>().Single();
        Assert.Equal(new DateOnly(2024, 3, 4), row.NullableDate);
        Assert.Equal(new TimeOnly(14, 30), row.NullableTime);
    }

    public class DateOnlyRow
    {
        public DateOnly Date { get; set; }
        public TimeOnly Time { get; set; }
        public DateOnly? NullableDate { get; set; }
        public TimeOnly? NullableTime { get; set; }
    }
#endif

    [Fact]
    public void ReadRows_WithValueFallback_UsedForEmptyAndInvalidCells()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n\nabc\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithFallback(5));
        Assert.Equal([1, 5, 5], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
    }

    public class Row
    {
        public DateTime DateTime { get; set; }
        public DateTimeOffset Offset { get; set; }
        public TimeSpan TimeSpan { get; set; }
    }

    public class NullableRow
    {
        public DateTime? DateTime { get; set; }
        public DateTimeOffset? Offset { get; set; }
        public TimeSpan? TimeSpan { get; set; }
    }

    public class ElementsRow
    {
        public DateTime[] Dates { get; set; } = [];
    }

    public class IntRow
    {
        public int Value { get; set; }
    }
}
