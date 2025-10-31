namespace ExcelMapper.Tests;

public class MapTimeSpanTests
{
    [Fact]
    public void ReadRow_TimeSpan_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TimeSpan>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeSpan>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeSpan>());
    }

    [Fact]
    public void ReadRow_DefaultMappedTimeSpan_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<TimeSpan>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TimeSpan>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeSpan>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeSpan>());
    }

    [Fact]
    public void ReadRow_CustomMappedTimeSpan_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<TimeSpan>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(new TimeSpan(02, 03, 04))
                .WithInvalidFallback(new TimeSpan(03, 04, 05));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TimeSpan>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<TimeSpan>();
        Assert.Equal(new TimeSpan(02, 03, 04), row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<TimeSpan>();
        Assert.Equal(new TimeSpan(03, 04, 05), row3);
    }

    [Fact]
    public void ReadRow_NullableTimeSpan_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TimeSpan?>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<TimeSpan?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeSpan?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableTimeSpan_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<TimeSpan?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TimeSpan?>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<TimeSpan?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeSpan?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableTimeSpan_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<TimeSpan?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(new TimeSpan(02, 03, 04))
                .WithInvalidFallback(new TimeSpan(03, 04, 05));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TimeSpan?>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<TimeSpan?>();
        Assert.Equal(new TimeSpan(02, 03, 04), row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<TimeSpan?>();
        Assert.Equal(new TimeSpan(03, 04, 05), row3);
    }

    [Fact]
    public void ReadRow_AutoMappedFormatsAttributeTimeSpan_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FormatsTimeSpanValue>();
        Assert.Equal(new TimeSpan(02, 01, 01), row1.CustomValue);

        var row2 = sheet.ReadRow<FormatsTimeSpanValue>();
        Assert.Equal(new TimeSpan(18, 30, 00), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsTimeSpanValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsTimeSpanValue>());
    }

    private class FormatsTimeSpanValue
    {
        [ExcelFormats("yyyy-MM-dd", "g")]
        public TimeSpan CustomValue { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedFormatsAttributeTimeSpan_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<FormatsTimeSpanValue>(c =>
        {
            c.Map(o => o.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FormatsTimeSpanValue>();
        Assert.Equal(new TimeSpan(02, 01, 01), row1.CustomValue);

        var row2 = sheet.ReadRow<FormatsTimeSpanValue>();
        Assert.Equal(new TimeSpan(18, 30, 00), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsTimeSpanValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsTimeSpanValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedTimeSpanValue_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TimeSpanValue>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeSpanValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeSpanValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedTimeSpanValue_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<TimeSpanValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TimeSpanValue>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeSpanValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeSpanValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedTimeSpanValue_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<TimeSpanValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new TimeSpan(02, 03, 04))
                .WithInvalidFallback(new TimeSpan(03, 04, 05));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<TimeSpanValue>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<TimeSpanValue>();
        Assert.Equal(new TimeSpan(02, 03, 04), row2.Value);

        // Invalid cell value.
        var row6 = sheet.ReadRow<TimeSpanValue>();
        Assert.Equal(new TimeSpan(03, 04, 05), row6.Value);
    }

    private class TimeSpanValue
    {
        public TimeSpan Value { get; set; }
    }

    [Fact]
    public void ReadRow_CustomFormatsArrayTimeSpan_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<CustomTimeSpanValue>(c =>
        {
            c.Map(o => o.CustomValue)
                .WithFormats("yyyy-MM-dd", "g");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomTimeSpanValue>();
        Assert.Equal(new TimeSpan(02, 01, 01), row1.CustomValue);

        var row2 = sheet.ReadRow<CustomTimeSpanValue>();
        Assert.Equal(new TimeSpan(18, 30, 00), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeSpanValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeSpanValue>());
    }

    [Fact]
    public void ReadRow_CustomFormatsEnumerableTimeSpan_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<CustomTimeSpanValue>(c =>
        {
            c.Map(o => o.CustomValue)
                .WithFormats((IEnumerable<string>)["yyyy-MM-dd", "g"]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomTimeSpanValue>();
        Assert.Equal(new TimeSpan(02, 01, 01), row1.CustomValue);

        var row2 = sheet.ReadRow<CustomTimeSpanValue>();
        Assert.Equal(new TimeSpan(18, 30, 00), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeSpanValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeSpanValue>());
    }

    private class CustomTimeSpanValue
    {
        public TimeSpan CustomValue { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedNullableTimeSpanValue_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableTimeSpanValue>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableTimeSpanValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableTimeSpanValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableTimeSpanValue_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<NullableTimeSpanValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableTimeSpanValue>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableTimeSpanValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableTimeSpanValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableTimeSpanValue_Success()
    {
        using var importer = Helpers.GetImporter("TimeSpans.xlsx");
        importer.Configuration.RegisterClassMap<NullableTimeSpanValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new TimeSpan(02, 03, 04))
                .WithInvalidFallback(new TimeSpan(03, 04, 05));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableTimeSpanValue>();
        Assert.Equal(new TimeSpan(01, 01, 01), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableTimeSpanValue>();
        Assert.Equal(new TimeSpan(02, 03, 04), row2.Value);

        // Invalid cell value.
        var row6 = sheet.ReadRow<NullableTimeSpanValue>();
        Assert.Equal(new TimeSpan(03, 04, 05), row6.Value);
    }

    private class NullableTimeSpanValue
    {
        public TimeSpan? Value { get; set; }
    }

/*

    [Fact]
    public void ReadRow_InvalidFormat_ThrowsErrorWithRightMessage()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using var importer = Helpers.GetImporter("TimeSpans_Errors.xlsx");
        importer.Configuration.RegisterClassMap<TimeSpanValueClassMap>();

        var sheet = importer.ReadSheet("Sheet1");
        sheet.HeadingIndex = 0;

        // Valid cell value.
        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<TimeSpanValueInt>(1, 1).ToArray());
        Assert.IsType<InvalidCastException>(ex.InnerException);
        Assert.Equal($"Cannot assign \"{new TimeSpan(2025, 01, 01, 1, 1, 1)}\" to member \"Value\" of type \"System.Int32\" in column \"Date\" on row 2 in sheet \"Sheet1\".", ex.Message);
    }

    private class TimeSpanValueInt
    {
        public int Value { get; set; }
    }

    private class TimeSpanValueClassMap : ExcelClassMap<TimeSpanValueInt>
    {
        public TimeSpanValueClassMap()
        {
            Map(m => m.Value)
                .WithColumnIndex(0);
        }
    }

    private class CustomTimeSpanValue
    {
        public TimeSpan Value { get; set; }
    }

    private class DefaultTimeSpanClassMap : ExcelClassMap<TimeSpanValue>
    {
        public DefaultTimeSpanClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomTimeSpanClassMap : ExcelClassMap<TimeSpanValue>
    {
        public CustomTimeSpanClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(new TimeSpan(2017, 07, 20))
                .WithInvalidFallback(new TimeSpan(2017, 07, 21));
        }
    }

    private class TimeSpanFormatsArrayMap : ExcelClassMap<CustomTimeSpanValue>
    {
        public TimeSpanFormatsArrayMap()
        {
            Map(o => o.Value)
                .WithFormats("yyyy-MM-dd", "G");
        }
    }

    private class TimeSpanEnumerableFormatsMap : ExcelClassMap<CustomTimeSpanValue>
    {
        public TimeSpanEnumerableFormatsMap()
        {
            Map(o => o.Value)
                .WithFormats(new List<string> { "yyyy-MM-dd", "G" });
        }
    }

    private class DefaultNullableTimeSpanClassMap : ExcelClassMap<NullableTimeSpanValue>
    {
        public DefaultNullableTimeSpanClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableTimeSpanClassMap : ExcelClassMap<NullableTimeSpanValue>
    {
        public CustomNullableTimeSpanClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(new TimeSpan(2017, 07, 20))
                .WithInvalidFallback(new TimeSpan(2017, 07, 21));
        }
    }
*/
}
