using System.Linq;
using System.Text;

namespace ExcelMapper.Tests;

public class MapDateTimeTests
{
    [Fact]
    public void ReadRow_DateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTime>();
        Assert.Equal(new DateTime(2017, 07, 19), row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTime>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTime>());
    }

    [Fact]
    public void ReadRow_NullableDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTime?>();
        Assert.Equal(new DateTime(2017, 07, 19), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateTime?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTime?>());
    }

    [Fact]
    public void ReadRow_AutoMappedFormatsAttributeDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FormatsDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.CustomValue);

        var row2 = sheet.ReadRow<FormatsDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 18), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateTimeValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateTimeValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedFormatsAttributeDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<FormatsDateTimeValue>(c =>
        {
            c.Map(o => o.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FormatsDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.CustomValue);

        var row2 = sheet.ReadRow<FormatsDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 18), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateTimeValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateTimeValue>());
    }

    private class FormatsDateTimeValue
    {
        [ExcelFormats("yyyy-MM-dd", "G")]
        public DateTime CustomValue { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());
    }

    private class DateTimeValue
    {
        public DateTime Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new DateTime(2017, 07, 20))
                .WithInvalidFallback(new DateTime(2017, 07, 21));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 20), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 21), row3.Value);
    }

    [Fact]
    public void ReadRow_CustomFormatsArrayDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<CustomDateTimeValue>(c =>
        {
            c.Map(o => o.CustomValue)
                .WithFormats("yyyy-MM-dd", "G");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.CustomValue);

        var row2 = sheet.ReadRow<CustomDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 18), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());
    }

    private class CustomDateTimeValue
    {
        public DateTime CustomValue { get; set; }
    }

    [Fact]
    public void ReadRow_CustomEnumerableFormatsDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<CustomDateTimeValue>(c =>
        {
            c.Map(o => o.CustomValue)
                .WithFormats(new List<string> { "yyyy-MM-dd", "G" });
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.CustomValue);

        var row2 = sheet.ReadRow<CustomDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 18), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDateTimeValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDateTimeValue>());
    }

    private class NullableDateTimeValue
    {
        public DateTime? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<NullableDateTimeValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDateTimeValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDateTimeValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDateTime_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<NullableDateTimeValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new DateTime(2017, 07, 20))
                .WithInvalidFallback(new DateTime(2017, 07, 21));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 20), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableDateTimeValue>();
        Assert.Equal(new DateTime(2017, 07, 21), row3.Value);
    }

    [Fact]
    public void ReadRow_InvalidFormat_ThrowsErrorWithRightMessage()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using var importer = Helpers.GetImporter("DateTimes_Errors.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeValueInt>(c =>
        {
            c.Map(o => o.Value)
                .WithColumnIndex(0);
        });

        var sheet = importer.ReadSheet("Sheet1");
        sheet.HeadingIndex = 0;

        // Valid cell value.
        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<DateTimeValueInt>(1, 1).ToArray());
        Assert.IsType<InvalidCastException>(ex.InnerException);
        Assert.Equal($"Cannot assign \"{new DateTime(2025, 01, 01, 1, 1, 1)}\" to member \"Value\" of type \"System.Int32\" in column \"Date\" on row 2 in sheet \"Sheet1\".", ex.Message);
    }

    private class DateTimeValueInt
    {
        public int Value { get; set; }
    }
}
