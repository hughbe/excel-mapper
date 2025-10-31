using System.Globalization;

namespace ExcelMapper.Tests;

public class MapUInt32Tests
{
    [Fact]
    public void ReadRow_UInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<uint>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint>());
    }

    [Fact]
    public void ReadRow_NullableUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<uint?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<uint?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<uint>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<uint>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint>());
    }

    [Fact]
    public void ReadRow_CustomMappedUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<uint>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(11u)
                .WithInvalidFallback(10u);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<uint>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<uint>();
        Assert.Equal(11u, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<uint>();
        Assert.Equal(10u, row3);
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<uint?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<uint?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<uint?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<uint?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback(11u)
                .WithInvalidFallback(10u);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<uint?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<uint?>();
        Assert.Equal(11u, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<uint?>();
        Assert.Equal(10u, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableUInt32Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<UInt32Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableUInt32Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableUInt32Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedUInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<UInt32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(11u)
                .WithInvalidFallback(10u);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt32Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableUInt32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(11u)
                .WithInvalidFallback(10u);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_UInt32Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint>());
    }

    private class UInt32Value
    {
        public uint Value { get; set; }
    }

    private class NullableUInt32Value
    {
        public uint? Value { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedUInt32ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(11u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(10u, row9.Value);
    }

    private class NumberStyleAttributeUInt32Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue(11u)]
        [ExcelInvalidValue(10u)]
        public uint Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt32ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeUInt32Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(11u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeUInt32Value>();
        Assert.Equal(10u, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedUInt32ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<UInt32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback(11u)
                .WithInvalidFallback(10u);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(11u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(10u, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedUInt32ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<UInt32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(11u)
                .WithInvalidFallback(10u);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(2345u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUInt32ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(11u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(10u, row9.Value);
    }

    private class NumberStyleAttributeNullableUInt32Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue(11u)]
        [ExcelInvalidValue(10u)]
        public uint? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt32ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableUInt32Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(11u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableUInt32Value>();
        Assert.Equal(10u, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt32ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NullableUInt32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback(11u)
                .WithInvalidFallback(10u);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(0xABu, row1.Value);

        var row2 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(0x123u, row2.Value);

        var row3 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(0xABu, row3.Value);

        var row4 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(0x123u, row4.Value);

        var row5 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(123u, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(11u, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(10u, row7.Value);

        var row8 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(10u, row8.Value);

        var row9 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(10u, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt32ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableUInt32Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback(11u)
                .WithInvalidFallback(10u);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(2345u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableUInt32Value>();
        Assert.Equal(10u, row3.Value);
    }
}
