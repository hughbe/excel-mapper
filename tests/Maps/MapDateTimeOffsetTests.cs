using System.Linq;
using System.Text;

namespace ExcelMapper.Tests;

public class MapDateTimeOffsetTests
{
    [Fact]
    public void ReadRow_DateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffset>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffset>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffset>());
    }

    [Fact]
    public void ReadRow_DefaultMappedDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeOffset>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffset>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffset>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffset>());
    }

    [Fact]
    public void ReadRow_CustomMappedDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeOffset>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(new DateTimeOffset(new DateTime(2017, 07, 20), TimeSpan.Zero))
                .WithInvalidFallback(new DateTimeOffset(new DateTime(2017, 07, 21), TimeSpan.Zero));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffset>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateTimeOffset>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 20), TimeSpan.Zero), row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DateTimeOffset>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 21), TimeSpan.Zero), row3);
    }

    [Fact]
    public void ReadRow_NullableDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffset?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeOffset?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffset?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeOffset?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(new DateTimeOffset(new DateTime(2017, 07, 20), TimeSpan.Zero))
                .WithInvalidFallback(new DateTimeOffset(new DateTime(2017, 07, 21), TimeSpan.Zero));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 20), TimeSpan.Zero), row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 21), TimeSpan.Zero), row3);
    }

    [Fact]
    public void ReadRow_AutoMappedFormatsAttributeDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FormatsDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1.CustomValue);

        var row2 = sheet.ReadRow<FormatsDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 18), TimeSpan.Zero), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateTimeOffsetValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateTimeOffsetValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedFormatsAttributeDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<FormatsDateTimeOffsetValue>(c =>
        {
            c.Map(o => o.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FormatsDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1.CustomValue);

        var row2 = sheet.ReadRow<FormatsDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 18), TimeSpan.Zero), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateTimeOffsetValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateTimeOffsetValue>());
    }

    private class FormatsDateTimeOffsetValue
    {
        [ExcelFormats("yyyy-MM-dd", "G")]
        public DateTimeOffset CustomValue { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedDateTimeOffsetValue_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffsetValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffsetValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedDateTimeOffsetValue_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeOffsetValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffsetValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffsetValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedDateTimeOffsetValue_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeOffsetValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new DateTimeOffset(new DateTime(2017, 07, 20), TimeSpan.Zero))
                .WithInvalidFallback(new DateTimeOffset(new DateTime(2017, 07, 21), TimeSpan.Zero));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 20), TimeSpan.Zero), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 21), TimeSpan.Zero), row3.Value);
    }

    private class DateTimeOffsetValue
    {
        public DateTimeOffset Value { get; set; }
    }

    [Fact]
    public void ReadRow_CustomFormatsArrayDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<CustomDateTimeOffsetValue>(c =>
        {
            c.Map(o => o.CustomValue)
                .WithFormats("yyyy-MM-dd", "G");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1.CustomValue);

        var row2 = sheet.ReadRow<CustomDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 18), TimeSpan.Zero), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffsetValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffsetValue>());
    }

    private class CustomDateTimeOffsetValue
    {
        public DateTimeOffset CustomValue { get; set; }
    }

    [Fact]
    public void ReadRow_CustomEnumerableFormatsDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<CustomDateTimeOffsetValue>(c =>
        {
            c.Map(o => o.CustomValue)
                .WithFormats(new List<string> { "yyyy-MM-dd", "G" });
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1.CustomValue);

        var row2 = sheet.ReadRow<CustomDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 18), TimeSpan.Zero), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffsetValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeOffsetValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableDateTimeOffsetValue_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDateTimeOffsetValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDateTimeOffsetValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDateTimeOffsetValue_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<NullableDateTimeOffsetValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDateTimeOffsetValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDateTimeOffsetValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDateTimeOffsetValue_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<NullableDateTimeOffsetValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new DateTimeOffset(new DateTime(2017, 07, 20), TimeSpan.Zero))
                .WithInvalidFallback(new DateTimeOffset(new DateTime(2017, 07, 21), TimeSpan.Zero));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19), TimeSpan.Zero), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 20), TimeSpan.Zero), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 21), TimeSpan.Zero), row3.Value);
    }

    private class NullableDateTimeOffsetValue
    {
        public DateTimeOffset? Value { get; set; }
    }

    [Fact]
    public void ReadRow_InvalidFormat_ThrowsErrorWithRightMessage()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using var importer = Helpers.GetImporter("DateTimes_Errors.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeOffsetValueInt>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnIndex(0);
        });

        var sheet = importer.ReadSheet("Sheet1");
        sheet.HeadingIndex = 0;

        // Valid cell value.
        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<DateTimeOffsetValueInt>(1, 1).ToArray());
#if LIBRARY_NET8
        // The inner exception depends on the library build (the netstandard builds use Convert.ChangeType).
        Assert.IsType<FormatException>(ex.InnerException);
#endif
        Assert.Equal($"Cannot assign \"{new DateTime(2025, 01, 01, 1, 1, 1)}\" to member \"Value\" of type int in column \"Date\" at cell A2 in sheet \"Sheet1\".", ex.Message);
    }

    private class DateTimeOffsetValueInt
    {
        public int Value { get; set; }
    }

    private static readonly DateTime s_offsetDate = new(2024, 3, 4, 5, 6, 7);

    public class OffsetRow
    {
        public DateTimeOffset When { get; set; }
    }

    public class OffsetNullableRow
    {
        public DateTimeOffset? When { get; set; }
    }

    public class OffsetUtcRow
    {
        [ExcelOffset("+00:00")]
        public DateTimeOffset When { get; set; }
    }

    public class OffsetNegativeRow
    {
        [ExcelOffset("-08:00")]
        public DateTimeOffset When { get; set; }
    }

    [Fact]
    public void ReadRows_DateCellDefaultOffset_UsesUtc()
    {
        // Not the machine's local offset, so every machine reads the same value.
        using var importer = XlsxBuilder.Create(["When"], [s_offsetDate]);
        var when = importer.ReadSheet().ReadRows<OffsetRow>().Single().When;
        Assert.Equal(s_offsetDate, when.DateTime);
        Assert.Equal(TimeSpan.Zero, when.Offset);
    }

    [Theory]
    [InlineData("2024-03-04T10:30:15", 10, 0)]
    [InlineData("2024-03-04 10:30", 10, 0)]
    [InlineData("03/04/2024 10:30:15", 10, 0)]
    [InlineData("2024-03-04T10:30:15Z", 10, 0)]
    [InlineData("2024-03-04T10:30:15+05:30", 10, 5.5)]
    [InlineData("2024-03-04T10:30:15-08:00", 10, -8)]
    public void ReadRows_TextDefaultOffset_UsesTextOffsetOrUtc(string text, int hour, double offsetHours)
    {
        using var importer = Helpers.GetCsvImporter($"When\n{text}\n");
        var when = importer.ReadSheet().ReadRows<OffsetRow>().Single().When;
        Assert.Equal(hour, when.Hour);
        Assert.Equal(TimeSpan.FromHours(offsetHours), when.Offset);
    }

    [Theory]
    [InlineData("2024-03-04T10:30:15", -8)]
    [InlineData("03/04/2024 10:30:15", -8)]
    [InlineData("2024-03-04T10:30:15Z", 0)]
    [InlineData("2024-03-04T10:30:15+05:30", 5.5)]
    public void ReadRows_TextWithMemberOffset_UsesTextOffsetOrMembers(string text, double offsetHours)
    {
        // The member's offset is for dates without one of their own.
        using var importer = Helpers.GetCsvImporter($"When\n{text}\n");
        var when = importer.ReadSheet().ReadRows<OffsetNegativeRow>().Single().When;
        Assert.Equal(10, when.Hour);
        Assert.Equal(TimeSpan.FromHours(offsetHours), when.Offset);
    }

    [Fact]
    public void ReadRows_InvalidTextWithMemberOffset_IsInvalid()
    {
        using var importer = Helpers.GetCsvImporter("When\nnot a date\n");
        var error = Assert.Single(Assert.Single(importer.ReadSheet().ReadRowResults<OffsetNegativeRow>()).Errors);
        Assert.IsType<FormatException>(error.InnerException);
    }

    [Fact]
    public void ReadRows_DateCellWithOffset_UsesOffset()
    {
        using var importer = XlsxBuilder.Create(["When"], [s_offsetDate]);
        importer.Configuration.RegisterClassMap<OffsetRow>(c => c.Map(p => p.When).WithOffset(TimeSpan.FromHours(5.5)));

        var when = importer.ReadSheet().ReadRows<OffsetRow>().Single().When;
        Assert.Equal(new DateTimeOffset(s_offsetDate, TimeSpan.FromHours(5.5)), when);
        Assert.Equal(TimeSpan.FromHours(5.5), when.Offset);
    }

    [Fact]
    public void ReadRows_NullableAndSerialNumberWithOffset_UsesOffset()
    {
        using var importer = XlsxBuilder.Create(["When"], [s_offsetDate.ToOADate()], [null]);
        importer.Configuration.RegisterClassMap<OffsetNullableRow>(c => c.Map(p => p.When).WithOffset(TimeSpan.Zero));

        var rows = importer.ReadSheet().ReadRows<OffsetNullableRow>().ToArray();
        Assert.Equal(new DateTimeOffset(s_offsetDate, TimeSpan.Zero), rows[0].When);
        Assert.Null(rows[1].When);
    }

    [Fact]
    public void ReadRows_ExcelOffsetAttribute_UsesOffset()
    {
        using var utc = XlsxBuilder.Create(["When"], [s_offsetDate]);
        Assert.Equal(new DateTimeOffset(s_offsetDate, TimeSpan.Zero), utc.ReadSheet().ReadRows<OffsetUtcRow>().Single().When);

        using var negative = XlsxBuilder.Create(["When"], [s_offsetDate]);
        Assert.Equal(new DateTimeOffset(s_offsetDate, TimeSpan.FromHours(-8)), negative.ReadSheet().ReadRows<OffsetNegativeRow>().Single().When);
    }

    [Fact]
    public void ReadRows_TextWithOffset_KeepsTextOffset()
    {
        using var importer = Helpers.GetCsvImporter("When\n2024-03-04T05:06:07+02:00\n");
        importer.Configuration.RegisterClassMap<OffsetRow>(c => c.Map(p => p.When).WithOffset(TimeSpan.Zero).WithFormats("yyyy-MM-ddTHH:mm:sszzz"));

        Assert.Equal(TimeSpan.FromHours(2), importer.ReadSheet().ReadRows<OffsetRow>().Single().When.Offset);
    }
}
