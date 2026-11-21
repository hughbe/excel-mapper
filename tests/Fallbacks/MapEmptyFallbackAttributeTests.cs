namespace ExcelMapper.Tests;

public class MapEmptyFallbackAttributeTests
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
        [ExcelEmptyFallback(typeof(CustomFallback), ConstructorArguments = [ "fallback" ])]
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
        [ExcelEmptyFallback(typeof(CustomFallback), ConstructorArguments = [ null ])]
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
        [ExcelEmptyFallback(typeof(CustomFallback), ConstructorArguments = [ 1 ])]
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

    private class CustomFallback(object? value) : ICellFallback<string?>, ICellFallback<int>
    {
        string? ICellFallback<string?>.Fallback(CellValue cell, CellContext context, Exception? exception) => (string?)value;

        int ICellFallback<int>.Fallback(CellValue cell, CellContext context, Exception? exception) => (int)value!;
    }
}
