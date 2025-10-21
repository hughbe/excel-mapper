using Xunit;

namespace ExcelMapper.Tests;

public class MapObjectTests
{
    [Fact]
    public void ReadRow_Object_Success()
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
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ObjectValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedObject_Success()
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

    [Fact]
    public void ReadRow_CustomMappedObject_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ObjectValueFallbackMap>();

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
        importer.Configuration.RegisterClassMap<RecordClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<RecordClass>();
        Assert.Equal(2, row1.Id!.Value);

        var row2 = sheet.ReadRow<RecordClass>();
        Assert.Null(row2.Id);

        var row3 = sheet.ReadRow<RecordClass>();
        Assert.Null(row3.Id);
    }

    private class ObjectValue
    {
        public object Value { get; set; } = default!;
    }

    private class ObjectValueFallbackMap : ExcelClassMap<ObjectValue>
    {
        public ObjectValueFallbackMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback("empty")
                .WithInvalidFallback("invalid");
        }
    }
    
    public record Id(int Value);

    public class RecordClass
    {
        public Id? Id { get; private set; }
    }

    public class RecordClassMap : ExcelClassMap<RecordClass>
    {
        public RecordClassMap()
        {
            Map(data => data.Id)
                .WithConverter(v => new Id(int.Parse(v!)))
                .WithColumnName("Value");
        }
    }
}
