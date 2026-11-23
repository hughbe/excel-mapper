using System.Linq;

namespace ExcelMapper.Tests;

public class MapInvalidValueAttributeTests
{
    [Fact]
    public void ReadRow_AutoMappedInvalidValueStringValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<InvalidValueStringClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<InvalidValueStringClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<InvalidValueStringClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<InvalidValueStringClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidValueStringClass>());
    }

    public class InvalidValueStringClass
    {
        [ExcelInvalidValue("fallback")]
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRow_InvalidMappedStringValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<InvalidValueStringClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<InvalidValueStringClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<InvalidValueStringClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<InvalidValueStringClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<InvalidValueStringClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidValueStringClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedInvalidValueStringNullValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<InvalidValueStringNullClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<InvalidValueStringNullClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<InvalidValueStringNullClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<InvalidValueStringNullClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidValueStringNullClass>());
    }

    public class InvalidValueStringNullClass
    {
        [ExcelInvalidValue(null)]
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRow_InvalidMappedStringNullValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<InvalidValueStringNullClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<InvalidValueStringNullClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<InvalidValueStringNullClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<InvalidValueStringNullClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<InvalidValueStringNullClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidValueStringNullClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedInvalidValueStringInvalidValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();

        // The value can't be the member's type, which is found when the class map is created.
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidValueStringInvalidClass>());
        Assert.Contains("given by ExcelInvalidValueAttribute on member \"Value\" can't be converted to", exception.Message);
    }

    public class InvalidValueStringInvalidClass
    {
        [ExcelInvalidValue(1)]
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRow_InvalidMappedStringInvalidValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        // The value can't be the member's type, which is found when the member is mapped.
        var exception = Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<InvalidValueStringInvalidClass>(c => c.Map(o => o.Value)));
        Assert.Contains("given by ExcelInvalidValueAttribute on member \"Value\" can't be converted to", exception.Message);
    }

    [Fact]
    public void ReadRow_AutoMappedIntInvalidValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<InvalidValueIntClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidValueIntClass>());

        // Invalid cell value.
        var row3 = sheet.ReadRow<InvalidValueIntClass>();
        Assert.Equal(1, row3.Value);
    }

    public class InvalidValueIntClass
    {
        [ExcelInvalidValue(1)]
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRow_AutoMappedIntEmptyNullValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();

        // The value can't be the member's type, which is found when the class map is created.
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidValueIntNullClass>());
        Assert.Equal("The value null given by ExcelInvalidValueAttribute on member \"Value\" can't be converted to int.", exception.Message);
    }

    public class InvalidValueIntNullClass
    {
        [ExcelInvalidValue(null)]
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRow_InvalidMappedInvalidValueInt_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<InvalidValueIntClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<InvalidValueIntClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidValueIntClass>());

        // Invalid cell value.
        var row3 = sheet.ReadRow<InvalidValueIntClass>();
        Assert.Equal(1, row3.Value);
    }

    [Fact]
    public void ReadRow_InvalidMappedEmptyNullValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        // The value can't be the member's type, which is found when the member is mapped.
        var exception = Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<InvalidValueIntNullClass>(c => c.Map(o => o.Value)));
        Assert.Equal("The value null given by ExcelInvalidValueAttribute on member \"Value\" can't be converted to int.", exception.Message);
    }

    [Fact]
    public void ReadRow_AutoMappedEmptyInvalidValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();

        // The value can't be the member's type, which is found when the class map is created.
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InvalidValueIntInvalidClass>());
        Assert.Contains("given by ExcelInvalidValueAttribute on member \"Value\" can't be converted to", exception.Message);
    }

    public class InvalidValueIntInvalidClass
    {
        [ExcelInvalidValue("fallback")]
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRow_InvalidMappedEmptyInvalidValue_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        // The value can't be the member's type, which is found when the member is mapped.
        var exception = Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<InvalidValueIntInvalidClass>(c => c.Map(o => o.Value)));
        Assert.Contains("given by ExcelInvalidValueAttribute on member \"Value\" can't be converted to", exception.Message);
    }

    [Fact]
    public void ReadRows_NumberInvalidValueOfAnotherNumberType_ReturnsConvertedValue()
    {
        using var importer = Helpers.GetCsvImporter("Value\ninvalid\n");
        Assert.Equal(-1m, importer.ReadSheet().ReadRows<DecimalInvalidValue>().Single().Value);
    }

    public class DecimalInvalidValue
    {
        [ExcelInvalidValue(-1)]
        public decimal Value { get; set; }
    }
}
