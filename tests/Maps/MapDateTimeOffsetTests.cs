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
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1);

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
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1);

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
                .WithEmptyFallback(new DateTimeOffset(new DateTime(2017, 07, 20)))
                .WithInvalidFallback(new DateTimeOffset(new DateTime(2017, 07, 21)));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffset>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateTimeOffset>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 20)), row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DateTimeOffset>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 21)), row3);
    }

    [Fact]
    public void ReadRow_NullableDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1);

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
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1);

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
                .WithEmptyFallback(new DateTimeOffset(new DateTime(2017, 07, 20)))
                .WithInvalidFallback(new DateTimeOffset(new DateTime(2017, 07, 21)));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 20)), row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DateTimeOffset?>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 21)), row3);
    }

    [Fact]
    public void ReadRow_AutoMappedFormatsAttributeDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FormatsDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1.CustomValue);

        var row2 = sheet.ReadRow<FormatsDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 18)), row2.CustomValue);

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
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1.CustomValue);

        var row2 = sheet.ReadRow<FormatsDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 18)), row2.CustomValue);

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
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1.Value);

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
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1.Value);

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
                .WithEmptyFallback(new DateTimeOffset(new DateTime(2017, 07, 20)))
                .WithInvalidFallback(new DateTimeOffset(new DateTime(2017, 07, 21)));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 20)), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 21)), row3.Value);
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
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1.CustomValue);

        var row2 = sheet.ReadRow<CustomDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 18)), row2.CustomValue);

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
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1.CustomValue);

        var row2 = sheet.ReadRow<CustomDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 18)), row2.CustomValue);

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
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1.Value);

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
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1.Value);

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
                .WithEmptyFallback(new DateTimeOffset(new DateTime(2017, 07, 20)))
                .WithInvalidFallback(new DateTimeOffset(new DateTime(2017, 07, 21)));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 19)), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 20)), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableDateTimeOffsetValue>();
        Assert.Equal(new DateTimeOffset(new DateTime(2017, 07, 21)), row3.Value);
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
        Assert.IsType<FormatException>(ex.InnerException);
        Assert.Equal($"Cannot assign \"{new DateTime(2025, 01, 01, 1, 1, 1)}\" to member \"Value\" of type \"System.Int32\" in column \"Date\" on row 1 in sheet \"Sheet1\".", ex.Message);
    }

    private class DateTimeOffsetValueInt
    {
        public int Value { get; set; }
    }
}
