using System.Globalization;

namespace ExcelMapper.Tests;

public class MapSByteTests
{
    [Fact]
    public void ReadRow_SByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<sbyte>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<sbyte>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<sbyte>());
    }

    [Fact]
    public void ReadRow_NullableSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<sbyte?>();
        Assert.Equal((sbyte)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<sbyte?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<sbyte?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<sbyte>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<sbyte>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<sbyte>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<sbyte>());
    }

    [Fact]
    public void ReadRow_CustomMappedSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<sbyte>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback((sbyte)-10)
                .WithInvalidFallback((sbyte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<sbyte>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<sbyte>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<sbyte>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<sbyte?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<sbyte?>();
        Assert.Equal((sbyte)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<sbyte?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<sbyte?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<sbyte?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback((sbyte)-10)
                .WithInvalidFallback((sbyte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<sbyte?>();
        Assert.Equal((sbyte)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<sbyte?>();
        Assert.Equal((sbyte)-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<sbyte?>();
        Assert.Equal((sbyte)10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedSByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<SByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableSByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableSByteValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableSByteValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedSByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<SByteValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<SByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<SByteValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableSByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableSByteValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableSByteValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableSByteValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedSByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<SByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((sbyte)-10)
                .WithInvalidFallback((sbyte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<SByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<SByteValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<SByteValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableSByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableSByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((sbyte)-10)
                .WithInvalidFallback((sbyte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)10, row3.Value);
    }

    [Fact]
    public void ReadRow_SByteOverflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<sbyte>());
    }

    private class SByteValue
    {
        public sbyte Value { get; set; }
    }

    private class NullableSByteValue
    {
        public sbyte? Value { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedSByteValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values - 0xAB (171) is -85 as signed byte
        var row1 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)-85, row1.Value); // 0xAB = 171 unsigned, -85 signed

        var row2 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)10, row2.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row3 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)-85, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)10, row4.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row5 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)10, row9.Value);
    }

    private class NumberStyleAttributeSByteValue
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue((sbyte)-10)]
        [ExcelInvalidValue((sbyte)10)]
        public sbyte Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedSByteValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeSByteValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values - 0xAB (171) is -85 as signed byte
        var row1 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)-85, row1.Value); // 0xAB = 171 unsigned, -85 signed

        var row2 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)10, row2.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row3 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)-85, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)10, row4.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row5 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeSByteValue>();
        Assert.Equal((sbyte)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedSByteValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<SByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback((sbyte)-10)
                .WithInvalidFallback((sbyte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values - 0xAB (171) is -85 as signed byte
        var row1 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)-85, row1.Value);

        var row2 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)10, row2.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row3 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)-85, row3.Value);

        var row4 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)10, row4.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row5 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)10, row7.Value);

        var row8 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)10, row8.Value);

        var row9 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedSByteValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<SByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback((sbyte)-10)
                .WithInvalidFallback((sbyte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)10, row1.Value); // 2345 overflows sbyte, uses invalid fallback

        // Empty cell value.
        var row2 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<SByteValue>();
        Assert.Equal((sbyte)10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableSByteValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values - 0xAB (171) is -85 as signed byte
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)-85, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)10, row2.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)-85, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)10, row4.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)10, row9.Value);
    }

    private class NumberStyleAttributeNullableSByteValue
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue((sbyte)-10)]
        [ExcelInvalidValue((sbyte)10)]
        public sbyte? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableSByteValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableSByteValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values - 0xAB (171) is -85 as signed byte
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)-85, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)10, row2.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)-85, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)10, row4.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableSByteValue>();
        Assert.Equal((sbyte)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableSByteValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NullableSByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback((sbyte)-10)
                .WithInvalidFallback((sbyte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values - 0xAB (171) is -85 as signed byte
        var row1 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)-85, row1.Value);

        var row2 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)10, row2.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row3 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)-85, row3.Value);

        var row4 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)10, row4.Value); // 0x123 overflows sbyte, uses invalid fallback

        var row5 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)10, row7.Value);

        var row8 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)10, row8.Value);

        var row9 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableSByteValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableSByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback((sbyte)-10)
                .WithInvalidFallback((sbyte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)10, row1.Value); // 2345 overflows sbyte, uses invalid fallback

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableSByteValue>();
        Assert.Equal((sbyte)10, row3.Value);
    }
}
