namespace ExcelMapper.Tests;

public class PreserveFormattingTests
{
    [Fact]
    public void ReadRow_LeadingZeroesInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<int>();
        Assert.Equal(123, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int>());
        
        // Default cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int>());
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesString_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<string>();
        Assert.Equal("123", row1);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<string>();
        Assert.Null(row2);
        
        // Default cell value.
        var row3 = sheet.ReadRow<string>();
        Assert.Equal("abc", row3);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesAutoMapped_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("123", row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<StringValue>();
        Assert.Null(row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<StringValue>();
        Assert.Equal("abc", row3.Value);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesDefaultMapped_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");
        importer.Configuration.RegisterClassMap<CustomStringValueClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<StringValue>();
        Assert.Equal("00123", row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<StringValue>();
        Assert.Null(row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<StringValue>();
        Assert.Equal("abc", row3.Value);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesAutoMappedAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<PreserveFormattingStringValue>();
        Assert.Equal("00123", row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<PreserveFormattingStringValue>();
        Assert.Null(row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<PreserveFormattingStringValue>();
        Assert.Equal("abc", row3.Value);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesAttributeDefaultMapped_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");
        importer.Configuration.RegisterClassMap<CustomStringValueClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<PreserveFormattingStringValue>();
        Assert.Equal("00123", row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<PreserveFormattingStringValue>();
        Assert.Null(row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<PreserveFormattingStringValue>();
        Assert.Equal("abc", row3.Value);
    }

    [Fact]
    public void ReadRow_LeadingZeroesInt32Enumerable_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<int[]>();
        Assert.Equal([123], row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int[]>());
        
        // Default cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int[]>());
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesStringEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<string?[]?>();
        Assert.Equal(new string[] { "123" }, row1);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<string?[]?>();
        Assert.Equal(new string?[] { null }, row2);
        
        // Default cell value.
        var row3 = sheet.ReadRow<string?[]?>();
        Assert.Equal(new string[] { "abc" }, row3);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesAutoMappedEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<StringsValue>();
        Assert.Equal(new string[] { "123" }, row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<StringsValue>();
        Assert.Empty(row2.Value!);
        
        // Default cell value.
        var row3 = sheet.ReadRow<StringsValue>();
        Assert.Equal(new string[] { "abc" }, row3.Value);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesDefaultMappedEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");
        importer.Configuration.RegisterClassMap<CustomStringsValueClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<StringsValue>();
        Assert.Equal(new string[] { "00123" }, row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<StringsValue>();
        Assert.Equal(new string?[] { null }, row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<StringsValue>();
        Assert.Equal(new string[] { "abc" }, row3.Value);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesAttributeAutoMappedEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<PreserveFormattingStringsValue>();
        Assert.Equal(new string[] { "00123" }, row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<PreserveFormattingStringsValue>();
        Assert.Equal(new string?[] { null }, row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<PreserveFormattingStringsValue>();
        Assert.Equal(new string[] { "abc" }, row3.Value);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesAttributeDefaultMappedEnumerable_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");
        importer.Configuration.RegisterClassMap<CustomStringsValueClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<PreserveFormattingStringsValue>();
        Assert.Equal(new string[] { "00123" }, row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<PreserveFormattingStringsValue>();
        Assert.Equal(new string?[] { null }, row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<PreserveFormattingStringsValue>();
        Assert.Equal(new string[] { "abc" }, row3.Value);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesInt32Dictionary_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<int[]>();
        Assert.Equal([123], row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int[]>());
        
        // Default cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<int[]>());
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesStringDictionary_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Dictionary<string, object?>>();
        Assert.Equal(new Dictionary<string, object?> { {"Value", "123" } }, row1);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<Dictionary<string, object?>>();
        Assert.Equal(new Dictionary<string, object?> { { "Value", null } }, row2);
        
        // Default cell value.
        var row3 = sheet.ReadRow<Dictionary<string, object?>>();
        Assert.Equal(new Dictionary<string, object?> { {"Value", "abc" } }, row3);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesAutoMappedDictionary_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { {"Value", "123" } }, row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<DictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { { "Value", null } }, row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<DictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { {"Value", "abc" } }, row3.Value);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesDefaultMappedDictionary_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryValueClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<DictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { {"Value", "00123" } }, row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<DictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { { "Value", null } }, row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<DictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { {"Value", "abc" } }, row3.Value);
    }
    
    
    [Fact]
    public void ReadRow_LeadingZeroesAttributeAutoMappedDictionary_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<PreserveFormattingDictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { {"Value", "00123" } }, row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<PreserveFormattingDictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { { "Value", null } }, row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<PreserveFormattingDictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { {"Value", "abc" } }, row3.Value);
    }
    
    [Fact]
    public void ReadRow_LeadingZeroesAttributeDefaultMappedDictionary_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_LeadingZeroes.xlsx");
        importer.Configuration.RegisterClassMap<CustomDictionaryValueClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<PreserveFormattingDictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { {"Value", "00123" } }, row1.Value);
        
        // Empty cell value.
        var row2 = sheet.ReadRow<PreserveFormattingDictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { { "Value", null } }, row2.Value);
        
        // Default cell value.
        var row3 = sheet.ReadRow<PreserveFormattingDictionaryValue>();
        Assert.Equal(new Dictionary<string, object?> { {"Value", "abc" } }, row3.Value);
    }

    public class StringValue
    {
        public string? Value { get; set; }
    }

    public class StringValueClassMap : ExcelClassMap<StringValue>
    {
        public StringValueClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class CustomStringValueClassMap : ExcelClassMap<StringValue>
    {
        public CustomStringValueClassMap()
        {
            Map(o => o.Value)
                .MakePreserveFormatting();
        }
    }

    public class PreserveFormattingStringValue
    {
        [ExcelPreserveFormatting]
        public string? Value { get; set; }
    }

    public class PreserveFormattingStringValueClassMap : ExcelClassMap<PreserveFormattingStringValue>
    {
        public PreserveFormattingStringValueClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class StringsValue
    {
        public string?[] Value { get; set; } = default!;
    }

    public class StringsValueClassMap : ExcelClassMap<StringsValue>
    {
        public StringsValueClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class CustomStringsValueClassMap : ExcelClassMap<StringsValue>
    {
        public CustomStringsValueClassMap()
        {
            Map(o => o.Value)
                .MakePreserveFormatting();
        }
    }

    public class PreserveFormattingStringsValue
    {
        [ExcelPreserveFormatting]
        public string?[] Value { get; set; } = default!;
    }

    public class PreserveFormattingStringsValueClassMap : ExcelClassMap<PreserveFormattingStringsValue>
    {
        public PreserveFormattingStringsValueClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class DictionaryValue
    {
        public Dictionary<string, object?> Value { get; set; } = default!;
    }

    public class DictionaryValueClassMap : ExcelClassMap<DictionaryValue>
    {
        public DictionaryValueClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class CustomDictionaryValueClassMap : ExcelClassMap<DictionaryValue>
    {
        public CustomDictionaryValueClassMap()
        {
            Map(o => o.Value)
                .MakePreserveFormatting();
        }
    }

    public class PreserveFormattingDictionaryValue
    {
        [ExcelPreserveFormatting]
        public Dictionary<string, object?> Value { get; set; } = default!;
    }

    public class PreserveFormattingDictionaryValueClassMap : ExcelClassMap<PreserveFormattingDictionaryValue>
    {
        public PreserveFormattingDictionaryValueClassMap()
        {
            Map(o => o.Value);
        }
    }
}