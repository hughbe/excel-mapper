namespace ExcelMapper.Tests;

public class MapTimeOnlyTests
{
    [Fact]
    public void ReadRow_TimeOnly_Success()
    {
        using var importer = Helpers.GetImporter("TimeOnly.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid time.
        var row1 = sheet.ReadRow<TimeOnly>();
        Assert.Equal(new TimeOnly(07, 01, 00), row1);

        // Earliest time.
        var row2 = sheet.ReadRow<TimeOnly>();
        Assert.Equal(new TimeOnly(00, 00, 00), row2);

        // Latest time.
        var row3 = sheet.ReadRow<TimeOnly>();
        Assert.Equal(new TimeOnly(23, 59, 59), row3);

        // Valid duration.
        var row4 = sheet.ReadRow<TimeOnly>();
        Assert.Equal(new TimeOnly(432620000000), row4);

        // Earliest duration.
        var row5 = sheet.ReadRow<TimeOnly>();
        Assert.Equal(new TimeOnly(00, 00, 00), row5);

        // Latest duration.
        var row6 = sheet.ReadRow<TimeOnly>();
        Assert.Equal(new TimeOnly(23, 59, 59), row6);

        // Large time.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnly>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnly>());

        // Negative cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnly>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnly>());
    }

    [Fact]
    public void ReadRow_AutoMappedTimeOnly_Success()
    {
        using var importer = Helpers.GetImporter("TimeOnly.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid time.
        var row1 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(07, 01, 00), row1.Value);

        // Earliest time.
        var row2 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row2.Value);

        // Latest time.
        var row3 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row3.Value);

        // Valid duration.
        var row4 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(432620000000), row4.Value);

        // Earliest duration.
        var row5 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row5.Value);

        // Latest duration.
        var row6 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row6.Value);

        // Large time.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnlyValue>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnlyValue>());

        // Negative cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnlyValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedTimeOnly_Success()
    {
        using var importer = Helpers.GetImporter("TimeOnly.xlsx");
        importer.Configuration.RegisterClassMap<TimeOnlyValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid time.
        var row1 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(07, 01, 00), row1.Value);

        // Earliest time.
        var row2 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row2.Value);

        // Latest time.
        var row3 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row3.Value);

        // Valid duration.
        var row4 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(432620000000), row4.Value);

        // Earliest duration.
        var row5 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row5.Value);

        // Latest duration.
        var row6 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row6.Value);

        // Large time.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnlyValue>());

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnlyValue>());

        // Negative cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnlyValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedTimeOnly_Success()
    {
        using var importer = Helpers.GetImporter("TimeOnly.xlsx");
        importer.Configuration.RegisterClassMap<TimeOnlyValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new TimeOnly(02, 03, 04))
                .WithInvalidFallback(new TimeOnly(03, 04, 05));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid time.
        var row1 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(07, 01, 00), row1.Value);

        // Earliest time.
        var row2 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row2.Value);

        // Latest time.
        var row3 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row3.Value);

        // Valid duration.
        var row4 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(432620000000), row4.Value);

        // Earliest duration.
        var row5 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row5.Value);

        // Latest duration.
        var row6 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row6.Value);

        // Large time.
        var row7 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(03, 04, 05), row7.Value);

        // Empty cell value.
        var row8 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(02, 03, 04), row8.Value);

        // Negative cell value.
        var row9 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(03, 04, 05), row9.Value);

        // Invalid cell value.
        var row10 = sheet.ReadRow<TimeOnlyValue>();
        Assert.Equal(new TimeOnly(03, 04, 05), row10.Value);
    }

    private class TimeOnlyValue
    {
        public TimeOnly Value { get; set; }
    }

    [Fact]
    public void ReadRow_NullableTimeOnly_Success()
    {
        using var importer = Helpers.GetImporter("TimeOnly.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid time.
        var row1 = sheet.ReadRow<TimeOnly?>();
        Assert.Equal(new TimeOnly(07, 01, 00), row1);

        // Earliest time.
        var row2 = sheet.ReadRow<TimeOnly?>();
        Assert.Equal(new TimeOnly(00, 00, 00), row2);

        // Latest time.
        var row3 = sheet.ReadRow<TimeOnly?>();
        Assert.Equal(new TimeOnly(23, 59, 59), row3);

        // Valid duration.
        var row4 = sheet.ReadRow<TimeOnly?>();
        Assert.Equal(new TimeOnly(432620000000), row4);

        // Earliest duration.
        var row5 = sheet.ReadRow<TimeOnly?>();
        Assert.Equal(new TimeOnly(00, 00, 00), row5);

        // Latest duration.
        var row6 = sheet.ReadRow<TimeOnly?>();
        Assert.Equal(new TimeOnly(23, 59, 59), row6);

        // Large time.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnly?>());

        // Empty cell value.
        var row8 = sheet.ReadRow<TimeOnly?>();
        Assert.Null(row8);

        // Negative cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnly?>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TimeOnly?>());
    }

    [Fact]
    public void ReadRow_CustomFormatsArrayTimeOnly_Success()
    {
        using var importer = Helpers.GetImporter("TimeOnly.xlsx");
        importer.Configuration.RegisterClassMap<CustomTimeOnlyValue>(c =>
        {
            c.Map(o => o.CustomValue)
                .WithFormats("yyyy-MM-dd", "o");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid time.
        var row1 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(07, 01, 00), row1.CustomValue);

        // Earliest time.
        var row2 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row2.CustomValue);

        // Latest time.
        var row3 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row3.CustomValue);

        // Valid duration.
        var row4 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(432620000000), row4.CustomValue);

        // Earliest duration.
        var row5 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row5.CustomValue);

        // Latest duration.
        var row6 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row6.CustomValue);

        // Large time.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeOnlyValue>());

        // Custom format.
        var row7 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(286660000000), row7.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeOnlyValue>());

        // Negative cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeOnlyValue>());
    }

    [Fact]
    public void ReadRow_CustomFormatsEnumerableTimeOnly_Success()
    {
        using var importer = Helpers.GetImporter("TimeOnly.xlsx");
        importer.Configuration.RegisterClassMap<CustomTimeOnlyValue>(c =>
        {
            c.Map(o => o.CustomValue)
                .WithFormats((IEnumerable<string>)["yyyy-MM-dd", "o"]);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid time.
        var row1 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(07, 01, 00), row1.CustomValue);

        // Earliest time.
        var row2 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row2.CustomValue);

        // Latest time.
        var row3 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row3.CustomValue);

        // Valid duration.
        var row4 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(432620000000), row4.CustomValue);

        // Earliest duration.
        var row5 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row5.CustomValue);

        // Latest duration.
        var row6 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row6.CustomValue);

        // Large time.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeOnlyValue>());

        // Custom format.
        var row7 = sheet.ReadRow<CustomTimeOnlyValue>();
        Assert.Equal(new TimeOnly(286660000000), row7.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeOnlyValue>());

        // Negative cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomTimeOnlyValue>());
    }

    private class CustomTimeOnlyValue
    {
        public TimeOnly CustomValue { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedNullableTimeOnly_Success()
    {
        using var importer = Helpers.GetImporter("TimeOnly.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid time.
        var row1 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(07, 01, 00), row1.Value);

        // Earliest time.
        var row2 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row2.Value);

        // Latest time.
        var row3 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row3.Value);

        // Valid duration.
        var row4 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(432620000000), row4.Value);

        // Earliest duration.
        var row5 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row5.Value);

        // Latest duration.
        var row6 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row6.Value);

        // Large time.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableTimeOnlyValue>());

        // Empty cell value.
        var row8 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Null(row8.Value);

        // Negative cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableTimeOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableTimeOnlyValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableTimeOnly_Success()
    {
        using var importer = Helpers.GetImporter("TimeOnly.xlsx");
        importer.Configuration.RegisterClassMap<NullableTimeOnlyValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid time.
        var row1 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(07, 01, 00), row1.Value);

        // Earliest time.
        var row2 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row2.Value);

        // Latest time.
        var row3 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row3.Value);

        // Valid duration.
        var row4 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(432620000000), row4.Value);

        // Earliest duration.
        var row5 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row5.Value);

        // Latest duration.
        var row6 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row6.Value);

        // Large time.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableTimeOnlyValue>());

        // Empty cell value.
        var row8 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Null(row8.Value);

        // Negative cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableTimeOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableTimeOnlyValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableTimeOnly_Success()
    {
        using var importer = Helpers.GetImporter("TimeOnly.xlsx");
        importer.Configuration.RegisterClassMap<NullableTimeOnlyValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new TimeOnly(02, 03, 04))
                .WithInvalidFallback(new TimeOnly(03, 04, 05));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid time.
        var row1 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(07, 01, 00), row1.Value);

        // Earliest time.
        var row2 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row2.Value);

        // Latest time.
        var row3 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row3.Value);

        // Valid duration.
        var row4 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(432620000000), row4.Value);

        // Earliest duration.
        var row5 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(00, 00, 00), row5.Value);

        // Latest duration.
        var row6 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(23, 59, 59), row6.Value);

        // Large time.
        var row7 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(03, 04, 05), row7.Value);

        // Empty cell value.
        var row8 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(02, 03, 04), row8.Value);

        // Negative cell value.
        var row9 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(03, 04, 05), row9.Value);

        // Invalid cell value.
        var row10 = sheet.ReadRow<NullableTimeOnlyValue>();
        Assert.Equal(new TimeOnly(03, 04, 05), row10.Value);
    }

    private class NullableTimeOnlyValue
    {
        public TimeOnly? Value { get; set; }
    }
}
