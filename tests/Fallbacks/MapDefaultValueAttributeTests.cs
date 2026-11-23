using System.Linq;

namespace ExcelMapper.Tests;

public class MapDefaultValueAttributeTests
{
    [Fact]
    public void ReadRow_AutoMappedDefaultValueStringValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("fallback", row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringClass>());
    }

    public class EmptyValueStringClass
    {
        [ExcelDefaultValue("fallback")]
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedStringValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueStringClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("fallback", row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedDefaultValueStringNullValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringNullClass>());
    }

    public class EmptyValueStringNullClass
    {
        [ExcelDefaultValue(null)]
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedStringNullValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueStringNullClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringNullClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedDefaultValueStringInvalidValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();

        // The value can't be the member's type, which is found when the class map is created.
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringInvalidClass>());
        Assert.Contains("given by ExcelDefaultValueAttribute on member \"Value\" can't be converted to", exception.Message);
    }

    public class EmptyValueStringInvalidClass
    {
        [ExcelDefaultValue(1)]
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedStringInvalidValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        // The value can't be the member's type, which is found when the member is mapped.
        var exception = Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<EmptyValueStringInvalidClass>(c => c.Map(o => o.Value)));
        Assert.Contains("given by ExcelDefaultValueAttribute on member \"Value\" can't be converted to", exception.Message);
    }

    [Fact]
    public void ReadRow_AutoMappedIntEmptyValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<EmptyValueIntClass>();
        Assert.Equal(1, row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntClass>());
    }

    public class EmptyValueIntClass
    {
        [ExcelDefaultValue(1)]
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedIntEmptyNullValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();

        // The value can't be the member's type, which is found when the class map is created.
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntNullClass>());
        Assert.Equal("The value null given by ExcelDefaultValueAttribute on member \"Value\" can't be converted to int.", exception.Message);
    }

    public class EmptyValueIntNullClass
    {
        [ExcelDefaultValue(null)]
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedEmptyValueInt_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueIntClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<EmptyValueIntClass>();
        Assert.Equal(1, row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedEmptyNullValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        // The value can't be the member's type, which is found when the member is mapped.
        var exception = Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<EmptyValueIntNullClass>(c => c.Map(o => o.Value)));
        Assert.Equal("The value null given by ExcelDefaultValueAttribute on member \"Value\" can't be converted to int.", exception.Message);
    }

    [Fact]
    public void ReadRow_AutoMappedEmptyInvalidValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();

        // The value can't be the member's type, which is found when the class map is created.
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntInvalidClass>());
        Assert.Contains("given by ExcelDefaultValueAttribute on member \"Value\" can't be converted to", exception.Message);
    }

    public class EmptyValueIntInvalidClass
    {
        [ExcelDefaultValue("fallback")]
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedEmptyInvalidValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        // The value can't be the member's type, which is found when the member is mapped.
        var exception = Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<EmptyValueIntInvalidClass>(c => c.Map(o => o.Value)));
        Assert.Contains("given by ExcelDefaultValueAttribute on member \"Value\" can't be converted to", exception.Message);
    }

    [Fact]
    public void ReadRows_NumberDefaultValueOfAnotherNumberType_ReturnsConvertedValue()
    {
        using var importer = Helpers.GetCsvImporter("Double,Decimal,Long,NullableShort,Color,Other\n,,,,,x\n");
        var row = importer.ReadSheet().ReadRows<NumberDefaults>().Single();
        Assert.Equal(0.5, row.Double);
        Assert.Equal(2.25m, row.Decimal);
        Assert.Equal(3L, row.Long);
        Assert.Equal((short?)4, row.NullableShort);
        Assert.Equal(ConsoleColor.Blue, row.Color);
    }

    public class NumberDefaults
    {
        [ExcelDefaultValue(0.5f)]
        public double Double { get; set; }

        [ExcelDefaultValue(2.25)]
        public decimal Decimal { get; set; }

        [ExcelDefaultValue(3)]
        public long Long { get; set; }

        [ExcelDefaultValue(4)]
        public short? NullableShort { get; set; }

        [ExcelDefaultValue(9)]
        public ConsoleColor Color { get; set; }

        public string? Other { get; set; }
    }

    [Theory]
    [InlineData(typeof(FractionForInt), "The value 1.5 given by ExcelDefaultValueAttribute on member \"Value\" can't be converted to int.")]
    [InlineData(typeof(TooLargeForByte), "The value 300 given by ExcelDefaultValueAttribute on member \"Value\" can't be converted to byte.")]
    [InlineData(typeof(TextForInt), "The value \"1\" given by ExcelDefaultValueAttribute on member \"Value\" can't be converted to int.")]
    [InlineData(typeof(FractionForEnum), "The value 1.5 given by ExcelDefaultValueAttribute on member \"Value\" can't be converted to ConsoleColor.")]
    public void ReadRows_DefaultValueChangedByConversion_ThrowsExcelMappingException(Type type, string message)
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => type == typeof(FractionForInt) ? sheet.ReadRows<FractionForInt>().ToArray()
            : type == typeof(TooLargeForByte) ? sheet.ReadRows<TooLargeForByte>().ToArray()
            : type == typeof(TextForInt) ? sheet.ReadRows<TextForInt>().ToArray()
            : (object)sheet.ReadRows<FractionForEnum>().ToArray());
        Assert.Equal(message, exception.Message);
    }

    public class FractionForInt
    {
        [ExcelDefaultValue(1.5)]
        public int Value { get; set; }
    }

    public class TooLargeForByte
    {
        [ExcelDefaultValue(300)]
        public byte Value { get; set; }
    }

    public class TextForInt
    {
        [ExcelDefaultValue("1")]
        public int Value { get; set; }
    }

    public class FractionForEnum
    {
        [ExcelDefaultValue(1.5)]
        public ConsoleColor Value { get; set; }
    }
}
