using System.Globalization;

namespace ExcelMapper.Tests;

public class MapUInt16Tests
{
    [Fact]
    public void ReadRow_UInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ushort>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort>());
    }

    [Fact]
    public void ReadRow_NullableUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ushort?>();
        Assert.Equal((ushort)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<ushort?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ushort>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ushort>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort>());
    }

    [Fact]
    public void ReadRow_CustomMappedUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ushort>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback((ushort)11)
                .WithInvalidFallback((ushort)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ushort>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<ushort>();
        Assert.Equal(11, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ushort>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ushort?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ushort?>();
        Assert.Equal((ushort)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<ushort?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ushort?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback((ushort)11)
                .WithInvalidFallback((ushort)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ushort?>();
        Assert.Equal((ushort)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<ushort?>();
        Assert.Equal((ushort)11, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ushort?>();
        Assert.Equal((ushort)10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableUInt16Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<UInt16Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableUInt16Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableUInt16Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedUInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<UInt16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((ushort)11)
                .WithInvalidFallback((ushort)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt16Value>();
        Assert.Equal(11, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<UInt16Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableUInt16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((ushort)11)
                .WithInvalidFallback((ushort)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)11, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)10, row3.Value);
    }

    [Fact]
    public void ReadRow_UInt16Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort>());
    }

    private class UInt16Value
    {
        public ushort Value { get; set; }
    }

    private class NullableUInt16Value
    {
        public ushort? Value { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedUInt16ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)10, row9.Value);
    }

    private class NumberStyleAttributeUInt16Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue((ushort)11)]
        [ExcelInvalidValue((ushort)10)]
        public ushort Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt16ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeUInt16Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeUInt16Value>();
        Assert.Equal((ushort)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedUInt16ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<UInt16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback((ushort)11)
                .WithInvalidFallback((ushort)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)0xAB, row1.Value);

        var row2 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)0x123, row2.Value);

        var row3 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)0xAB, row3.Value);

        var row4 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)0x123, row4.Value);

        var row5 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)10, row7.Value);

        var row8 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)10, row8.Value);

        var row9 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedUInt16ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<UInt16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback((ushort)11)
                .WithInvalidFallback((ushort)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)11, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<UInt16Value>();
        Assert.Equal((ushort)10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUInt16ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)10, row9.Value);
    }

    private class NumberStyleAttributeNullableUInt16Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue((ushort)11)]
        [ExcelInvalidValue((ushort)10)]
        public ushort? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt16ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableUInt16Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableUInt16Value>();
        Assert.Equal((ushort)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt16ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NullableUInt16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback((ushort)11)
                .WithInvalidFallback((ushort)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)0x123, row2.Value);

        var row3 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)0x123, row4.Value);

        var row5 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)10, row7.Value);

        var row8 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)10, row8.Value);

        var row9 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt16ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableUInt16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback((ushort)11)
                .WithInvalidFallback((ushort)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)11, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableUInt16Value>();
        Assert.Equal((ushort)10, row3.Value);
    }
}
