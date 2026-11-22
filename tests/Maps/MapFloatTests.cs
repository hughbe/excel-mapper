using System.Globalization;

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
    public void ReadRow_DefaultMappedFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<float>(c =>
        {
            c.Map(p => p);
        });

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
    public void ReadRow_CustomMappedFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<float>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<float>();
        Assert.Equal(2.2345f, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<float>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<float>();
        Assert.Equal(10, row3);
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
    public void ReadRow_DefaultMappedNullableFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<float?>(c =>
        {
            c.Map(p => p);
        });

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
    public void ReadRow_CustomMappedNullableFloat_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<float?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<float?>();
        Assert.Equal(2.2345f, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<float?>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<float?>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedFloatValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FloatValue>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedFloatValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<FloatValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FloatValue>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedFloatValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeFloatValue>();
        Assert.Equal(2345.67f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeFloatValue>();
        Assert.Equal(-10.0f, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NumberStyleAttributeFloatValue>();
        Assert.Equal(10.0f, row3.Value);
    }

    private class NumberStyleAttributeFloatValue
    {
        [ExcelNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)]
        [ExcelDefaultValue(-10.0f)]
        [ExcelInvalidValue(10.0f)]
        public float Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedFloatValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeFloatValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeFloatValue>();
        Assert.Equal(2345.67f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeFloatValue>();
        Assert.Equal(-10.0f, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NumberStyleAttributeFloatValue>();
        Assert.Equal(10.0f, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedFloatValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<FloatValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FloatValue>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<FloatValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<FloatValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedFloatValueAllowThousands_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<FloatValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FloatValue>();
        Assert.Equal(2345.67f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<FloatValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<FloatValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedFloatValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<FloatValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<FloatValue>();
        Assert.Equal(2345.67f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<FloatValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<FloatValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableFloatValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableFloatValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableFloatValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableFloatValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableFloatValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableFloatValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableFloatValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableFloatValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableFloatValue>();
        Assert.Equal(2345.67f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeNullableFloatValue>();
        Assert.Equal(-10.0f, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NumberStyleAttributeNullableFloatValue>();
        Assert.Equal(10.0f, row3.Value);
    }

    private class NumberStyleAttributeNullableFloatValue
    {
        [ExcelNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)]
        [ExcelDefaultValue(-10.0f)]
        [ExcelInvalidValue(10.0f)]
        public float? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableFloatValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableFloatValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableFloatValue>();
        Assert.Equal(2345.67f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeNullableFloatValue>();
        Assert.Equal(-10.0f, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NumberStyleAttributeNullableFloatValue>();
        Assert.Equal(10.0f, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableFloatValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableFloatValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(2.2345f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableFloatValueAllowThousands_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NullableFloatValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(2345.67f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableFloatValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableFloatValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(-10.0f)
                .WithInvalidFallback(10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(2345.67f, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableFloatValue>();
        Assert.Equal(10, row3.Value);
    }

    private class FloatValue
    {
        public float Value { get; set; }
    }

    private class NullableFloatValue
    {
        public float? Value { get; set; }
    }
}
