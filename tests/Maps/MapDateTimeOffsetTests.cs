using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xunit;

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
    public void ReadRow_AutoMappedDateTimeOffset_Success()
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
    public void ReadRow_AutoMappedNullableDateTimeOffset_Success()
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
    public void ReadRow_DefaultMappedDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DefaultDateTimeOffsetClassMap>();

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
    public void ReadRow_DefaultMappedNullableDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableDateTimeOffsetClassMap>();

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
    public void ReadRow_CustomMappedDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<CustomDateTimeOffsetClassMap>();

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

    [Fact]
    public void ReadRow_CustomFormatsArrayDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeOffsetFormatsArrayMap>();

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
    public void ReadRow_CustomEnumerableFormatsDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeOffsetEnumerableFormatsMap>();

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
    public void ReadRow_CustomMappedNullableDateTimeOffset_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableDateTimeOffsetClassMap>();

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

    [Fact]
    public void ReadRow_InvalidFormat_ThrowsErrorWithRightMessage()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using var importer = Helpers.GetImporter("DateTimes_Errors.xlsx");
        importer.Configuration.RegisterClassMap<DateTimeOffsetValueClassMap>();

        var sheet = importer.ReadSheet("Sheet1");
        sheet.HeadingIndex = 0;

        // Valid cell value.
        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<DateTimeOffsetValueInt>(1, 1).ToArray());
        Assert.IsType<InvalidCastException>(ex.InnerException);
        Assert.Equal($"Cannot assign \"{new DateTime(2025, 01, 01, 1, 1, 1)}\" to member \"Value\" of type \"System.Int32\" in column \"Date\" on row 2 in sheet \"Sheet1\".", ex.Message);
    }

    private class DateTimeOffsetValueInt
    {
        public int Value { get; set; }
    }

    private class DateTimeOffsetValueClassMap : ExcelClassMap<DateTimeOffsetValueInt>
    {
        public DateTimeOffsetValueClassMap()
        {
            Map(m => m.Value)
                .WithColumnIndex(0);
        }
    }

    private class DateTimeOffsetValue
    {
        public DateTimeOffset Value { get; set; }
    }

    private class CustomDateTimeOffsetValue
    {
        public DateTimeOffset CustomValue { get; set; }
    }

    private class NullableDateTimeOffsetValue
    {
        public DateTimeOffset? Value { get; set; }
    }

    private class DefaultDateTimeOffsetClassMap : ExcelClassMap<DateTimeOffsetValue>
    {
        public DefaultDateTimeOffsetClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomDateTimeOffsetClassMap : ExcelClassMap<DateTimeOffsetValue>
    {
        public CustomDateTimeOffsetClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(new DateTimeOffset(new DateTime(2017, 07, 20)))
                .WithInvalidFallback(new DateTimeOffset(new DateTime(2017, 07, 21)));
        }
    }

    private class DateTimeOffsetFormatsArrayMap : ExcelClassMap<CustomDateTimeOffsetValue>
    {
        public DateTimeOffsetFormatsArrayMap()
        {
            Map(o => o.CustomValue)
                .WithFormats("yyyy-MM-dd", "G");
        }
    }

    private class DateTimeOffsetEnumerableFormatsMap : ExcelClassMap<CustomDateTimeOffsetValue>
    {
        public DateTimeOffsetEnumerableFormatsMap()
        {
            Map(o => o.CustomValue)
                .WithFormats(new List<string> { "yyyy-MM-dd", "G" });
        }
    }

    private class DefaultNullableDateTimeOffsetClassMap : ExcelClassMap<NullableDateTimeOffsetValue>
    {
        public DefaultNullableDateTimeOffsetClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableDateTimeOffsetClassMap : ExcelClassMap<NullableDateTimeOffsetValue>
    {
        public CustomNullableDateTimeOffsetClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(new DateTimeOffset(new DateTime(2017, 07, 20)))
                .WithInvalidFallback(new DateTimeOffset(new DateTime(2017, 07, 21)));
        }
    }
}
