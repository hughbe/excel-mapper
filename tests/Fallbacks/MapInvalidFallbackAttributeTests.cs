namespace ExcelMapper.Tests;

public class MapInvalidFallbackAttributeTests
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
        [ExcelInvalidFallback(typeof(CustomFallback), ConstructorArguments = ["fallback"])]
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
        [ExcelInvalidFallback(typeof(CustomFallback), ConstructorArguments = [null])]
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
        [ExcelInvalidFallback(typeof(CustomFallback), ConstructorArguments = [1])]
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

    private class CustomFallback(object? value) : ICellFallback<string?>, ICellFallback<int>
    {
        string? ICellFallback<string?>.Fallback(CellValue cell, CellContext context, Exception? exception) => (string?)value;

        int ICellFallback<int>.Fallback(CellValue cell, CellContext context, Exception? exception) => (int)value!;
    }
}
