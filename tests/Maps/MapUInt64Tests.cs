using System.Globalization;

namespace ExcelMapper.Tests;

public class MapUInt64Tests
{
    [Fact]
    public void ReadRow_UInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ulong>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ulong>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ulong>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong>());
    }

    [Fact]
    public void ReadRow_CustomMappedUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ulong>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10ul);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ulong>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<ulong>();
        Assert.Equal(11u, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ulong>();
        Assert.Equal(10u, row3);
    }

    [Fact]
    public void ReadRow_NullableUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ulong?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<ulong?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ulong?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ulong?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<ulong?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ulong?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10ul);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ulong?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<ulong?>();
        Assert.Equal(11u, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ulong?>();
        Assert.Equal(10u, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedUInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<UInt64Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedUInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<UInt64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10ul);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedUInt64ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<UInt64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback(10ul)
                .WithInvalidFallback(10ul);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(10u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(10u, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedUInt64ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<UInt64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(10ul)
                .WithInvalidFallback(10ul);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(2345u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(10u, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableUInt64Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableUInt64Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableUInt64Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt64Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableUInt64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((ulong)11)
                .WithInvalidFallback((ulong)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt64ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NullableUInt64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback(10ul)
                .WithInvalidFallback(10ul);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(10u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(10u, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt64ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableUInt64Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(10ul)
                .WithInvalidFallback(10ul);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(2345u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(10u, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableUInt64Value>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_UInt64Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong>());
    }

    private class UInt64Value
    {
        public ulong Value { get; set; }
    }

    private class NullableUInt64Value
    {
        public ulong? Value { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedUInt64ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(10u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(10u, row9.Value);
    }

    private class NumberStyleAttributeUInt64Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue(10ul)]
        [ExcelInvalidValue(10ul)]
        public ulong Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt64ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeUInt64Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(10u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeUInt64Value>();
        Assert.Equal(10u, row9.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUInt64ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(10u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(10u, row9.Value);
    }

    private class NumberStyleAttributeNullableUInt64Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue(10ul)]
        [ExcelInvalidValue(10ul)]
        public ulong? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt64ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableUInt64Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(10u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableUInt64Value>();
        Assert.Equal(10u, row9.Value);
    }
}
