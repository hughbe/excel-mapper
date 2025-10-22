namespace ExcelMapper.Tests;

public class MapFloatTests
{
    [Fact]
    public void ReadRow_Float_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<float>();
        Assert.Equal(2.2345f, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<float>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<float>());
    }

    [Fact]
    public void ReadRow_NullableFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<float?>();
        Assert.Equal(2.2345f, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<float?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<float?>());
    }

    [Fact]
    public void ReadRow_AutoMappedFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FloatClass>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatClass>());
    }

    private class FloatClass
    {
        public float Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<FloatClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FloatClass>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<FloatClass>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FloatClass>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<FloatClass>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<FloatClass>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableFloatClass>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableFloatClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableFloatClass>());
    }

    private class NullableFloatClass
    {
        public float? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableFloatClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableFloatClass>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableFloatClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableFloatClass>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableFloatClass>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableFloatClass>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableFloatClass>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableFloatClass>();
        Assert.Equal(10, row3.Value);
    }
}
