using System.Globalization;

namespace ExcelMapper.Tests;

public class MapInt16Tests
{
    [Fact]
    public void ReadRow_AutoMappedInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<short>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<short>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<short>());
    }

    [Fact]
    public void ReadRow_DefaultMappedInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<short>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<short>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<short>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<short>());
    }

    [Fact]
    public void ReadRow_CustomMappedInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<short>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback((short)-10)
                .WithInvalidFallback((short)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<short>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<short>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<short>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<short?>();
        Assert.Equal((short)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<short?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<short?>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<short?>(c =>
        {
            c.Map(p => p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<short?>();
        Assert.Equal((short)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<short?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<short?>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<short?>(c =>
        {
            c.Map(p => p)
                .WithEmptyFallback((short)-10)
                .WithInvalidFallback((short)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<short?>();
        Assert.Equal((short)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<short?>();
        Assert.Equal((short)-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<short?>();
        Assert.Equal((short)10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());
    }

    private class Int16Value
    {
        public short Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<Int16Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<Int16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((short)-10)
                .WithInvalidFallback((short)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<Int16Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<Int16Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedInt16ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)10, row9.Value);
    }

    private class NumberStyleAttributeInt16Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue((short)-10)]
        [ExcelInvalidValue((short)10)]
        public short Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedInt16ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeInt16Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeInt16Value>();
        Assert.Equal((short)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedInt16ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<Int16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback((short)-10)
                .WithInvalidFallback((short)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)0xAB, row1.Value);

        var row2 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)0x123, row2.Value);

        var row3 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)0xAB, row3.Value);

        var row4 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)0x123, row4.Value);

        var row5 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)10, row7.Value);

        var row8 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)10, row8.Value);

        var row9 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedInt16ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<Int16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback((short)-10)
                .WithInvalidFallback((short)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<Int16Value>();
        Assert.Equal((short)10, row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableInt16ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)10, row9.Value);
    }

    private class NumberStyleAttributeNullableInt16Value
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        [ExcelDefaultValue((short)-10)]
        [ExcelInvalidValue((short)10)]
        public short? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt16ValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableInt16Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)10, row7.Value);

        var row8 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)10, row8.Value);

        var row9 = sheet.ReadRow<NumberStyleAttributeNullableInt16Value>();
        Assert.Equal((short)10, row9.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedNullableInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt16Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());
    }

    private class NullableInt16Value
    {
        public short? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableInt16Value>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt16Value>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int16Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt16Value_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<NullableInt16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback((short)-10)
                .WithInvalidFallback((short)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt16ValueHex_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NullableInt16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithNumberStyle(NumberStyles.HexNumber)
                .WithEmptyFallback((short)-10)
                .WithInvalidFallback((short)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)0x123, row2.Value);

        var row3 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)0x123, row4.Value);

        var row5 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)-10, row6.Value);

        // Invalid cell values.
        var row7 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)10, row7.Value);

        var row8 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)10, row8.Value);

        var row9 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)10, row9.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt16ValueFormatProvider_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_FormatProvider.xlsx");
        var customFormatProvider = new NumberFormatInfo
        {
            NumberGroupSeparator = ";"
        };
        importer.Configuration.RegisterClassMap<NullableInt16Value>(c =>
        {
            c.Map(o => o.Value)
                .WithFormatProvider(customFormatProvider)
                .WithEmptyFallback((short)-10)
                .WithInvalidFallback((short)10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)2345, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)-10, row2.Value);

        // Invalid cell values.
        var row3 = sheet.ReadRow<NullableInt16Value>();
        Assert.Equal((short)10, row3.Value);
    }

    [Fact]
    public void ReadRow_Int16Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<short>());
    }
}
