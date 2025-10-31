using System.Globalization;

namespace ExcelMapper.Tests;

public class MapDoubleTests
{
    [Fact]
    public void ReadRow_Double_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<double>();
        Assert.Equal(2.2345, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<double>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<double>());
    }

    [Fact]
    public void ReadRow_DefaultMappedDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<double>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<double>();
        Assert.Equal(2.2345, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<double>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<double>());
    }

    [Fact]
    public void ReadRow_CustomMappedDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<double>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(-10.0)
                .WithInvalidFallback(10.0);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<double>();
        Assert.Equal(2.2345, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<double>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<double>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_NullableDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<double?>();
        Assert.Equal(2.2345, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<double?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<double?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<double?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<double?>();
        Assert.Equal(2.2345, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<double?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<double?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDouble_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<double?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(-10.0)
                .WithInvalidFallback(10.0);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<double?>();
        Assert.Equal(2.2345, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<double?>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<double?>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedDoubleValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleValue>());
    }

    private class DoubleValue
    {
        public double Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDoubleValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<DoubleValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedDoubleValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeDoubleValue>();
        Assert.Equal(2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeDoubleValue>();
        Assert.Equal(-10.0, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NumberStyleAttributeDoubleValue>();
        Assert.Equal(10.0, row3.Value);
    }

    private class NumberStyleAttributeDoubleValue
    {
        [ExcelNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)]
        [ExcelDefaultValue(-10.0)]
        [ExcelInvalidValue(10.0)]
        public double Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDoubleValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeDoubleValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeDoubleValue>();
        Assert.Equal(2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeDoubleValue>();
        Assert.Equal(-10.0, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NumberStyleAttributeDoubleValue>();
        Assert.Equal(10.0, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedDoubleValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<DoubleValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0)
                .WithInvalidFallback(10.0);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedDoubleValueAllowThousands_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<DoubleValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithEmptyFallback(-10.0)
                .WithInvalidFallback(10.0);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedDoubleValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<DoubleValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(-10.0)
                .WithInvalidFallback(10.0);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<DoubleValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableDoubleValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDoubleValue>());
    }

    private class NullableDoubleValue
    {
        public double? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDoubleValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableDoubleValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDoubleValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableDoubleValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableDoubleValue>();
        Assert.Equal(2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeNullableDoubleValue>();
        Assert.Equal(-10.0, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NumberStyleAttributeNullableDoubleValue>();
        Assert.Equal(10.0, row3.Value);
    }

    private class NumberStyleAttributeNullableDoubleValue
    {
        [ExcelNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)]
        [ExcelDefaultValue(-10.0)]
        [ExcelInvalidValue(10.0)]
        public double? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDoubleValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableDoubleValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableDoubleValue>();
        Assert.Equal(2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeNullableDoubleValue>();
        Assert.Equal(-10.0, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NumberStyleAttributeNullableDoubleValue>();
        Assert.Equal(10.0, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDoubleValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableDoubleValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0)
                .WithInvalidFallback(10.0);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDoubleValueAllowThousands_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NullableDoubleValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithEmptyFallback(-10.0)
                .WithInvalidFallback(10.0);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDoubleValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableDoubleValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(-10.0)
                .WithInvalidFallback(10.0);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableDoubleValue>();
        Assert.Equal(10, row3.Value);
    }
}
