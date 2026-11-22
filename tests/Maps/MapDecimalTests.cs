using System.Globalization;

namespace ExcelMapper.Tests;

public class MapDecimalTests
{
    [Fact]
    public void ReadRow_Decimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<decimal>();
        Assert.Equal(2.2345m, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<decimal>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<decimal>());
    }

    [Fact]
    public void ReadRow_DefaultMappedDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<decimal>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<decimal>();
        Assert.Equal(2.2345m, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<decimal>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<decimal>());
    }

    [Fact]
    public void ReadRow_CustomMappedDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<decimal>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(-10.0m)
                .WithInvalidFallback(10.0m);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<decimal>();
        Assert.Equal(2.2345m, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<decimal>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<decimal>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_NullableDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<decimal?>();
        Assert.Equal(2.2345m, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<decimal?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<decimal?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<decimal?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<decimal?>();
        Assert.Equal(2.2345m, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<decimal?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<decimal?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDecimal_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<decimal?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(-10.0m)
                .WithInvalidFallback(10.0m);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<decimal?>();
        Assert.Equal(2.2345m, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<decimal?>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<decimal?>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedDecimalValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedDecimalValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<DecimalValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedDecimalValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeDecimalValue>();
        Assert.Equal(2345.67m, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeDecimalValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeDecimalValue>());
    }

    private class NumberStyleAttributeDecimalValue
    {
        [ExcelNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)]
        public decimal Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDecimalValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeDecimalValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeDecimalValue>();
        Assert.Equal(2345.67m, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeDecimalValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeDecimalValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedDecimalValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<DecimalValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0m)
                .WithInvalidFallback(10.0m);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedDecimalValueAllowThousands_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<DecimalValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithEmptyFallback(-10.0m)
                .WithInvalidFallback(10.0m);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(2345.67m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedDecimalValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<DecimalValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(-10.0m)
                .WithInvalidFallback(10.0m);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(2345.67m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<DecimalValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableDecimalValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDecimalValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDecimalValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableDecimalValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDecimalValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableDecimalValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableDecimalValue>();
        Assert.Equal(2345.67m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeNullableDecimalValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableDecimalValue>());
    }

    private class NumberStyleAttributeNullableDecimalValue
    {
        [ExcelNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)]
        public decimal? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableDecimalValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableDecimalValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableDecimalValue>();
        Assert.Equal(2345.67m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NumberStyleAttributeNullableDecimalValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableDecimalValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDecimalValue_Success()
    {
        using var importer = Helpers.GetImporter("Doubles.xlsx");
        importer.Configuration.RegisterClassMap<NullableDecimalValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10.0m)
                .WithInvalidFallback(10.0m);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(2.2345m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDecimalValueAllowThousands_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_Thousands.xlsx");
        importer.Configuration.RegisterClassMap<NullableDecimalValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithEmptyFallback(-10.0m)
                .WithInvalidFallback(10.0m);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(2345.67m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableDecimalValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Doubles_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableDecimalValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(-10.0m)
                .WithInvalidFallback(10.0m);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(2345.67m, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableDecimalValue>();
        Assert.Equal(10, row3.Value);
    }

    private class DecimalValue
    {
        public decimal Value { get; set; }
    }

    private class NullableDecimalValue
    {
        public decimal? Value { get; set; }
    }
}
