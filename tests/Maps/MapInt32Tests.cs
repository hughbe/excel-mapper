using System.Globalization;

namespace ExcelMapper.Tests;

public class MapInt32Tests
{
    [Fact]
    public void ReadRow_AutoMappedInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<int>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int>());
    }

    [Fact]
    public void ReadRow_DefaultMappedInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<int>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<int>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int>());
    }

    [Fact]
    public void ReadRow_CustomMappedInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<int>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<int>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<int>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<int>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<int?>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<int?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<int?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<int?>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<int?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<int?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<int?>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<int?>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<int?>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int32Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());
    }

    private class Int32Value
    {
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedInt32ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(10, row9.Value);
    }

    private class NumberStyleAttributeInt32Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue(-10)]
        [ExcelInvalidValue(10)]
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedInt32ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeInt32Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeInt32Value>();
        Assert.Equal(10, row9.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableInt32ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(10, row9.Value);
    }

    private class NumberStyleAttributeNullableInt32Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue(-10)]
        [ExcelInvalidValue(10)]
        public int? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt32ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableInt32Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableInt32Value>();
        Assert.Equal(10, row9.Value);
    }

    [Fact]
    public void ReadRow_DefaultMappedInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<Int32Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int32Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<Int32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int32Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<Int32Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<Int32Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedInt32ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<Int32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<Int32Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<Int32Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<Int32Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<Int32Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<Int32Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<Int32Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<Int32Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<Int32Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<Int32Value>();
        Assert.Equal(10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedInt32ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<Int32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int32Value>();
        Assert.Equal(2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<Int32Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<Int32Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt32Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());
    }

    private class NullableInt32Value
    {
        public int? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableInt32Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt32Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableInt32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt32ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NullableInt32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt32ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableInt32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableInt32Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_Int32Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int>());
    }
}
