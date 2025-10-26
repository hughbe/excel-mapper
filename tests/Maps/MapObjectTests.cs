namespace ExcelMapper.Tests;

public class MapObjectTests
{
    [Fact]
    public void ReadRow_AutoMappedObject_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<object>();
        Assert.Equal("value", row1);

        // Valid value
        var row2 = sheet.ReadRow<object>();
        Assert.Equal("  value  ", row2);

        // Empty value
        var row3 = sheet.ReadRow<object>();
        Assert.Null(row3);

        // Valid value
        var row4 = sheet.ReadRow<object>();
        Assert.Equal("value", row4);

        // No more rows
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<object>());
    }
    
    [Fact]
    public void ReadRow_DefaultMappedObject_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<object>(c =>
        {
            c.Map(o => o);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<object>();
        Assert.Equal("value", row1);

        // Valid value
        var row2 = sheet.ReadRow<object>();
        Assert.Equal("  value  ", row2);

        // Empty value
        var row3 = sheet.ReadRow<object>();
        Assert.Null(row3);

        // Valid value
        var row4 = sheet.ReadRow<object>();
        Assert.Equal("value", row4);
        
        // No more rows
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<object>());
    }
    
    [Fact]
    public void ReadRow_CustomMappedObject_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<object>(c =>
        {
            c.Map(o => o)
                .WithEmptyFallback("empty")
                .WithInvalidFallback("invalid");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<object>();
        Assert.Equal("value", row1);

        // Valid value
        var row2 = sheet.ReadRow<object>();
        Assert.Equal("  value  ", row2);

        // Empty value
        var row3 = sheet.ReadRow<object>();
        Assert.Equal("empty", row3);

        // Valid value
        var row4 = sheet.ReadRow<object>();
        Assert.Equal("value", row4);
        
        // No more rows
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<object>());
    }

    [Fact]
    public void ReadRow_DefaultMappedCast_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<object>(c =>
        {
            c.Map(p => (int)p);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<object>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<object>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<object>());
    }

    [Fact]
    public void ReadRow_CustomMappedCast_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<object>(c =>
        {
            c.Map(p => (int)p)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<object>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<object>();
        Assert.Equal(-10, row2);

        // Invalid cell value.
        var row3 = sheet.ReadRow<object>();
        Assert.Equal(10, row3);
    }

    [Fact]
    public void ReadRow_AutoMappedObjectValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ObjectValue>();
        Assert.Null(row3.Value);

        // Valid value
        var row4 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("value", row4.Value);

        // No more rows
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ObjectValue>());
    }

    private class ObjectValue
    {
        public object Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedObjectValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ObjectValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ObjectValue>();
        Assert.Null(row3.Value);

        // Valid value
        var row4 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("value", row4.Value);

        // No more rows
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ObjectValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedObjectValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ObjectValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback("empty")
                .WithInvalidFallback("invalid");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("empty", row3.Value);

        // Valid value
        var row4 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("value", row4.Value);
        
        // No more rows
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ObjectValue>());
    }

    [Fact]
    public void ReadRow_Record_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<RecordClass>(c =>
        {
            c.Map(data => data.Id)
                .WithConverter(v => new Id(int.Parse(v!)))
                .WithColumnName("Value");
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<RecordClass>();
        Assert.Equal(2, row1.Id!.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<RecordClass>();
        Assert.Null(row2.Id);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<RecordClass>());
    }
    
    public record Id(int Value);

    public class RecordClass
    {
        public Id? Id { get; private set; }
    }
}
