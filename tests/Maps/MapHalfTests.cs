using System.Globalization;

namespace ExcelMapper.Tests;

public class MapHalfTests
{
    [Fact]
    public void ReadRow_Half_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Half>();
        Assert.Equal((Half)2.2345, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Half>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Half>());
    }

    [Fact]
    public void ReadRow_DefaultMappedHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<Half>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Half>();
        Assert.Equal((Half)2.2345, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Half>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Half>());
    }

    [Fact]
    public void ReadRow_CustomMappedHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<Half>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback((Half)(-10.0f))
                .WithInvalidFallback((Half)10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Half>();
        Assert.Equal((Half)2.2345, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<Half>();
        Assert.Equal((Half)(-10), row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<Half>();
        Assert.Equal((Half)10, row3);
    }

    [Fact]
    public void ReadRow_NullableHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Half?>();
        Assert.Equal((Half)2.2345, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<Half?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Half?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<Half?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Half?>();
        Assert.Equal((Half)2.2345, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<Half?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Half?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableHalf_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<Half?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback((Half)(-10.0f))
                .WithInvalidFallback((Half)10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Half?>();
        Assert.Equal((Half)2.2345, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<Half?>();
        Assert.Equal((Half)(-10), row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<Half?>();
        Assert.Equal((Half)10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedHalfValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HalfValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HalfValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedHalfValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<HalfValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HalfValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<HalfValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedHalfValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeHalfValue>();
        Assert.Equal((Half)2345.67, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeHalfValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeHalfValue>());
    }

    private class NumberStyleAttributeHalfValue
    {
        [ExcelNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)]
        public Half Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedHalfValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeHalfValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeHalfValue>();
        Assert.Equal((Half)2345.67, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeHalfValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeHalfValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedHalfValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<HalfValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((Half)(-10.0f))
                .WithInvalidFallback((Half)10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)(-10), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedHalfValueAllowThousands_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<HalfValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithEmptyFallback((Half)(-10.0f))
                .WithInvalidFallback((Half)10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)(-10), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedHalfValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<HalfValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback((Half)(-10.0f))
                .WithInvalidFallback((Half)10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)(-10), row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<HalfValue>();
        Assert.Equal((Half)10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableHalfValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableHalfValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableHalfValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableHalfValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableHalfValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableHalfValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableHalfValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableHalfValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableHalfValue>();
        Assert.Equal((Half)2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeNullableHalfValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableHalfValue>());
    }

    private class NumberStyleAttributeNullableHalfValue
    {
        [ExcelNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)]
        public Half? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableHalfValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableHalfValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableHalfValue>();
        Assert.Equal((Half)2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeNullableHalfValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableHalfValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableHalfValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableHalfValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((Half)(-10.0f))
                .WithInvalidFallback((Half)10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)2.2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)(-10), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableHalfValueAllowThousands_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NullableHalfValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithEmptyFallback((Half)(-10.0f))
                .WithInvalidFallback((Half)10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)(-10), row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableHalfValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableHalfValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback((Half)(-10.0f))
                .WithInvalidFallback((Half)10.0f);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)2345.67, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)(-10), row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableHalfValue>();
        Assert.Equal((Half)10, row3.Value);
    }

    private class HalfValue
    {
        public Half Value { get; set; }
    }

    private class NullableHalfValue
    {
        public Half? Value { get; set; }
    }
}
