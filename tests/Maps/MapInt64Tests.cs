using System.Globalization;

namespace ExcelMapper.Tests;

public class MapInt64Tests
{
    [Fact]
    public void ReadRow_Int64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<long>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<long>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<long>());
    }

    [Fact]
    public void ReadRow_NullableInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<long?>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<long?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<long?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<long>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<long>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<long>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<long>());
    }

    [Fact]
    public void ReadRow_CustomMappedInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<long>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(-10L)
                .WithInvalidFallback(10L);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<long>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<long>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<long>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<long?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<long?>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<long?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<long?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<long?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(-10L)
                .WithInvalidFallback(10L);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<long?>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<long?>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<long?>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int64Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int64Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int64Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt64Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableInt64Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<Int64Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int64Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int64Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int64Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableInt64Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt64Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableInt64Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<Int64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10L)
                .WithInvalidFallback(10L);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int64Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<Int64Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<Int64Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableInt64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(-10L)
                .WithInvalidFallback(10L);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_Int64Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<long>());
    }

    private class Int64Value
    {
        public long Value { get; set; }
    }

    private class NullableInt64Value
    {
        public long? Value { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedInt64ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(10, row9.Value);
    }

    private class NumberStyleAttributeInt64Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue(-10L)]
        [ExcelInvalidValue(10L)]
        public long Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedInt64ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeInt64Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeInt64Value>();
        Assert.Equal(10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedInt64ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<Int64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback(-10L)
                .WithInvalidFallback(10L);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<Int64Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<Int64Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<Int64Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<Int64Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<Int64Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<Int64Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<Int64Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<Int64Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<Int64Value>();
        Assert.Equal(10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedInt64ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<Int64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(-10L)
                .WithInvalidFallback(10L);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int64Value>();
        Assert.Equal(2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<Int64Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<Int64Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableInt64ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(10, row9.Value);
    }

    private class NumberStyleAttributeNullableInt64Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue(-10L)]
        [ExcelInvalidValue(10L)]
        public long? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt64ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableInt64Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableInt64Value>();
        Assert.Equal(10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt64ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NullableInt64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback(-10L)
                .WithInvalidFallback(10L);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(0xAB, row1.Value);

        var row2 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(0x123, row2.Value);

        var row3 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(0xAB, row3.Value);

        var row4 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(0x123, row4.Value);

        var row5 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(10, row7.Value);

        var row8 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(10, row8.Value);

        var row9 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt64ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableInt64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(-10L)
                .WithInvalidFallback(10L);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableInt64Value>();
        Assert.Equal(10, row3.Value);
    }
}
