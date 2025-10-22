using System.Reflection;
using ExcelMapper.Abstractions;

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
    public void ReadRow_AutoMappedDefaultValueStringInvalidValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        Assert.Throws<InvalidCastException>(() => sheet.ReadRow<EmptyValueStringInvalidClass>());

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringInvalidClass>());
    }

    public class EmptyValueStringInvalidClass
    {
        [ExcelEmptyFallback(typeof(CustomFallback), ConstructorArguments = [ 1 ])]
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedStringInvalidValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueStringInvalidClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        Assert.Throws<InvalidCastException>(() => sheet.ReadRow<EmptyValueStringInvalidClass>());

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringInvalidClass>());
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
    public void ReadRow_AutoMappedIntEmptyNullValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntNullClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<NullReferenceException>(() => sheet.ReadRow<EmptyValueIntNullClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntNullClass>());
    }

    public class EmptyValueIntNullClass
    {
        [ExcelEmptyFallback(typeof(CustomFallback), ConstructorArguments = [ null ])]
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
    public void ReadRow_DefaultMappedEmptyNullValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueIntNullClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntNullClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<NullReferenceException>(() => sheet.ReadRow<EmptyValueIntNullClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntNullClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedEmptyInvalidValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntInvalidClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<InvalidCastException>(() => sheet.ReadRow<EmptyValueIntInvalidClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntInvalidClass>());
    }

    public class EmptyValueIntInvalidClass
    {
        [ExcelEmptyFallback(typeof(CustomFallback), ConstructorArguments = [ "fallback" ])]
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedEmptyInvalidValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueIntInvalidClass>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntInvalidClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<InvalidCastException>(() => sheet.ReadRow<EmptyValueIntInvalidClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntInvalidClass>());
    }

    private class CustomFallback : IFallbackItem
    {
        private object? _value;

        public CustomFallback(object? value)
        {
            _value = value;
        }

        public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
            => _value;
    }
}
