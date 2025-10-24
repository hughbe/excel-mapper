namespace ExcelMapper.Tests;

public class MapDateOnlyTests
{
    [Fact]
    public void ReadRow_DateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateOnly>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnly>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnly>());
    }

    [Fact]
    public void ReadRow_NullableDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateOnly?>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateOnly?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnly?>());
    }

    [Fact]
    public void ReadRow_AutoMappedFormatsAttributeDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FormatsDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1.CustomValue);

        var row2 = sheet.ReadRow<FormatsDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 18), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateOnlyValue>());
    }

    private class FormatsDateOnlyValue
    {
        [ExcelFormats("yyyy-MM-dd", "G")]
        public DateOnly CustomValue { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedFormatsAttributeDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");
        importer.Configuration.RegisterClassMap<FormatsDateOnlyValue>(c =>
        {
            c.Map(o => o.CustomValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FormatsDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1.CustomValue);

        var row2 = sheet.ReadRow<FormatsDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 18), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FormatsDateOnlyValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnlyValue>());
    }

    private class DateOnlyValue
    {
        public DateOnly Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");
        importer.Configuration.RegisterClassMap<DateOnlyValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnlyValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");
        importer.Configuration.RegisterClassMap<DateOnlyValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new DateOnly(2017, 07, 20))
                .WithInvalidFallback(new DateOnly(2017, 07, 21));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 20), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 21), row3.Value);
    }

    [Fact]
    public void ReadRow_CustomFormatsArrayDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");
        importer.Configuration.RegisterClassMap<CustomDateOnlyValue>(c =>
        {
            c.Map(o => o.CustomValue)
                .WithFormats("yyyy-MM-dd", "G");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1.CustomValue);

        var row2 = sheet.ReadRow<CustomDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 18), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnlyValue>());
    }

    private class CustomDateOnlyValue
    {
        public DateOnly CustomValue { get; set; }
    }

    [Fact]
    public void ReadRow_CustomEnumerableFormatsDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");
        importer.Configuration.RegisterClassMap<CustomDateOnlyValue>(c =>
        {
            c.Map(o => o.CustomValue)
                .WithFormats(new List<string> { "yyyy-MM-dd", "G" });
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<CustomDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1.CustomValue);

        var row2 = sheet.ReadRow<CustomDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 18), row2.CustomValue);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnlyValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateOnlyValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDateOnlyValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDateOnlyValue>());
    }

    private class NullableDateOnlyValue
    {
        public DateOnly? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");
        importer.Configuration.RegisterClassMap<NullableDateOnlyValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDateOnlyValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDateOnlyValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDateOnly_Success()
    {
        using var importer = Helpers.GetImporter("DateOnly.xlsx");
        importer.Configuration.RegisterClassMap<NullableDateOnlyValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new DateOnly(2017, 07, 20))
                .WithInvalidFallback(new DateOnly(2017, 07, 21));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 19), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 20), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableDateOnlyValue>();
        Assert.Equal(new DateOnly(2017, 07, 21), row3.Value);
    }
}
