namespace ExcelMapper.Tests;

public class MapIConvertibleTests
{
    [Fact]
    public void ReadRow_IConvertible_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<IConvertible>();
        Assert.Equal("value", row1);

        // Valid value
        var row2 = sheet.ReadRow<IConvertible>();
        Assert.Equal("  value  ", row2);

        // Empty value
        var row3 = sheet.ReadRow<IConvertible>();
        Assert.Null(row3);
    }

    [Fact]
    public void ReadRow_IConvertibleClass_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConvertibleClass>());

        // Valid value
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConvertibleClass>());

        // Empty value
        var row3 = sheet.ReadRow<ConvertibleClass>();
        Assert.Null(row3);
    }

    private class ConvertibleClass : IConvertible
    {
        public TypeCode GetTypeCode() => throw new NotImplementedException();

        public bool ToBoolean(IFormatProvider? provider) => throw new NotImplementedException();

        public byte ToByte(IFormatProvider? provider) => throw new NotImplementedException();

        public char ToChar(IFormatProvider? provider) => throw new NotImplementedException();

        public DateTime ToDateTime(IFormatProvider? provider) => throw new NotImplementedException();

        public decimal ToDecimal(IFormatProvider? provider) => throw new NotImplementedException();

        public double ToDouble(IFormatProvider? provider) => throw new NotImplementedException();

        public short ToInt16(IFormatProvider? provider) => throw new NotImplementedException();

        public int ToInt32(IFormatProvider? provider) => throw new NotImplementedException();

        public long ToInt64(IFormatProvider? provider) => throw new NotImplementedException();

        public sbyte ToSByte(IFormatProvider? provider) => throw new NotImplementedException();

        public float ToSingle(IFormatProvider? provider) => throw new NotImplementedException();

        public string ToString(IFormatProvider? provider) => throw new NotImplementedException();

        public object ToType(Type conversionType, IFormatProvider? provider)
        {
            throw new NotImplementedException();
        }

        public ushort ToUInt16(IFormatProvider? provider) => throw new NotImplementedException();

        public uint ToUInt32(IFormatProvider? provider) => throw new NotImplementedException();

        public ulong ToUInt64(IFormatProvider? provider) => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_IConvertibleStruct_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConvertibleStruct>());

        // Valid value
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConvertibleStruct>());

        // Empty value
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConvertibleStruct>());
    }

    private struct ConvertibleStruct : IConvertible
    {
        public TypeCode GetTypeCode() => throw new NotImplementedException();

        public bool ToBoolean(IFormatProvider? provider) => throw new NotImplementedException();

        public byte ToByte(IFormatProvider? provider) => throw new NotImplementedException();

        public char ToChar(IFormatProvider? provider) => throw new NotImplementedException();

        public DateTime ToDateTime(IFormatProvider? provider) => throw new NotImplementedException();

        public decimal ToDecimal(IFormatProvider? provider) => throw new NotImplementedException();

        public double ToDouble(IFormatProvider? provider) => throw new NotImplementedException();

        public short ToInt16(IFormatProvider? provider) => throw new NotImplementedException();

        public int ToInt32(IFormatProvider? provider) => throw new NotImplementedException();

        public long ToInt64(IFormatProvider? provider) => throw new NotImplementedException();

        public sbyte ToSByte(IFormatProvider? provider) => throw new NotImplementedException();

        public float ToSingle(IFormatProvider? provider) => throw new NotImplementedException();

        public string ToString(IFormatProvider? provider) => throw new NotImplementedException();

        public object ToType(Type conversionType, IFormatProvider? provider)
        {
            throw new NotImplementedException();
        }

        public ushort ToUInt16(IFormatProvider? provider) => throw new NotImplementedException();

        public uint ToUInt32(IFormatProvider? provider) => throw new NotImplementedException();

        public ulong ToUInt64(IFormatProvider? provider) => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_IConvertibleNullableStruct_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConvertibleStruct?>());

        // Valid value
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ConvertibleStruct?>());

        // Empty value
        var row3 = sheet.ReadRow<ConvertibleStruct?>();
        Assert.Null(row3);
    }

    [Fact]
    public void ReadRow_AutoMappedIConvertibleClassValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ConvertibleClassValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ConvertibleClassValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ConvertibleClassValue>();
        Assert.Null(row3.Value);
    }

    private class ConvertibleClassValue
    {
        public IConvertible Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedIConvertibleClassValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ConvertibleClassValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ConvertibleClassValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ConvertibleClassValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ConvertibleClassValue>();
        Assert.Null(row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIConvertibleClassValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ConvertibleClassValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback("empty")
                .WithInvalidFallback("invalid");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ConvertibleClassValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ConvertibleClassValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ConvertibleClassValue>();
        Assert.Equal("empty", row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedIConvertibleStructValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ConvertibleStructValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ConvertibleStructValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ConvertibleStructValue>();
        Assert.Null(row3.Value);
    }

    private struct ConvertibleStructValue
    {
        public ConvertibleStructValue()
        {
        }

        public IConvertible Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedIConvertibleStructValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ConvertibleStructValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ConvertibleStructValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ConvertibleStructValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ConvertibleStructValue>();
        Assert.Null(row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIConvertibleStructValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ConvertibleStructValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback("empty")
                .WithInvalidFallback("invalid");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ConvertibleStructValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ConvertibleStructValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ConvertibleStructValue>();
        Assert.Equal("empty", row3.Value);
    }
}
