using System.Globalization;

namespace ExcelMapper.Tests;

public class MapByteTests
{
    [Fact]
    public void ReadRow_Byte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<byte>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte>());
    }

    [Fact]
    public void ReadRow_NullableByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<byte?>();
        Assert.Equal((byte)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<byte?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<byte>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<byte>();
        Assert.Equal((byte)2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte>());
    }

    [Fact]
    public void ReadRow_CustomMappedByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<byte>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback((byte)11)
                .WithInvalidFallback((byte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<byte>();
        Assert.Equal((byte)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<byte>();
        Assert.Equal((byte)11, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<byte>();
        Assert.Equal((byte)10, row3);
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<byte?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<byte?>();
        Assert.Equal((byte)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<byte?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableByte_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<byte?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback((byte)11)
                .WithInvalidFallback((byte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<byte?>();
        Assert.Equal((byte)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<byte?>();
        Assert.Equal((byte)11, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<byte?>();
        Assert.Equal((byte)10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ByteValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ByteValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<ByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((byte)11)
                .WithInvalidFallback((byte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ByteValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ByteValue>();
        Assert.Equal(11, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ByteValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableByteValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableByteValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableByteValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableByteValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableByteValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableByteValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((byte)11)
                .WithInvalidFallback((byte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)11, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)10, row3.Value);
    }

    [Fact]
    public void ReadRow_ByteOverflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<byte>());
    }

    private class ByteValue
    {
        public byte Value { get; set; }
    }

    private class NullableByteValue
    {
        public byte? Value { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedByteValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)10, row2.Value); // 0x123 overflows byte, uses invalid fallback

        var row3 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)10, row4.Value); // 0x123 overflows byte, uses invalid fallback

        var row5 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)10, row9.Value);
    }

    private class NumberStyleAttributeByteValue
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue((byte)11)]
        [ExcelInvalidValue((byte)10)]
        public byte Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedByteValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeByteValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)10, row2.Value); // 0x123 overflows byte, uses invalid fallback

        var row3 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)10, row4.Value); // 0x123 overflows byte, uses invalid fallback

        var row5 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeByteValue>();
        Assert.Equal((byte)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedByteValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<ByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback((byte)11)
                .WithInvalidFallback((byte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)0xAB, row1.Value);

        var row2 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)10, row2.Value); // 0x123 overflows byte, uses invalid fallback

        var row3 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)0xAB, row3.Value);

        var row4 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)10, row4.Value); // 0x123 overflows byte, uses invalid fallback

        var row5 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)10, row7.Value);

        var row8 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)10, row8.Value);

        var row9 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedByteValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<ByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback((byte)11)
                .WithInvalidFallback((byte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)10, row1.Value); // 2345 overflows byte, uses invalid fallback

        // Empty cell value.
        var row2 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)11, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<ByteValue>();
        Assert.Equal((byte)10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableByteValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)10, row2.Value); // 0x123 overflows byte, uses invalid fallback

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)10, row4.Value); // 0x123 overflows byte, uses invalid fallback

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)10, row9.Value);
    }

    private class NumberStyleAttributeNullableByteValue
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue((byte)11)]
        [ExcelInvalidValue((byte)10)]
        public byte? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableByteValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableByteValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)10, row2.Value); // 0x123 overflows byte, uses invalid fallback

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)10, row4.Value); // 0x123 overflows byte, uses invalid fallback

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableByteValue>();
        Assert.Equal((byte)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableByteValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NullableByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback((byte)11)
                .WithInvalidFallback((byte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)10, row2.Value); // 0x123 overflows byte, uses invalid fallback

        var row3 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)10, row4.Value); // 0x123 overflows byte, uses invalid fallback

        var row5 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)11, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)10, row7.Value);

        var row8 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)10, row8.Value);

        var row9 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableByteValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableByteValue>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback((byte)11)
                .WithInvalidFallback((byte)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)10, row1.Value); // 2345 overflows byte, uses invalid fallback

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)11, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableByteValue>();
        Assert.Equal((byte)10, row3.Value);
    }
}
