using ExcelMapper.Abstractions;
using ExcelMapper.Mappers;

namespace ExcelMapper.Tests;

public class MapMapperAttributeTests
{
    [Fact]
    public void ReadRow_AutoMappedDateTimeMapper_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ObjectValue>());
    }

    private class ObjectValue
    {
        [ExcelMapper(typeof(DateTimeMapper))]
        public object? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedDateTimeMapper_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<ObjectValue>(c =>
        {
            c.Map(m => m.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectValue>();
        Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ObjectValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedDateTimeMapper_Success()
    {
        using var importer = Helpers.GetImporter("DateTimes.xlsx");
        importer.Configuration.RegisterClassMap<ObjectValue>(c =>
        {
            c.Map(m => m.Value)
                .WithMappers(new StringMapper());
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ObjectValue>();
        Assert.Equal(new DateTime(2017, 07, 19).ToString(), row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<ObjectValue>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<ObjectValue>();
        Assert.Equal("invalid", row3.Value);
    }

    [Fact]
    public void ReadRow_AutoMappedMultipleMappers_Success()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryClass>(c =>
        {
            c.Map(m => m.StringValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryClass>();
        Assert.Equal("12345", row1.StringValue);

        var row2 = sheet.ReadRow<DictionaryClass>();
        Assert.Equal("b", row2.StringValue);

        var row3 = sheet.ReadRow<DictionaryClass>();
        Assert.Equal("B", row3.StringValue);

        var row4 = sheet.ReadRow<DictionaryClass>();
        Assert.Null(row4.StringValue);
    }

    [Fact]
    public void ReadRow_DefaultMappedMultipleMappers_Success()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryClass>(c =>
        {
            c.Map(m => m.StringValue);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryClass>();
        Assert.Equal("12345", row1.StringValue);

        var row2 = sheet.ReadRow<DictionaryClass>();
        Assert.Equal("b", row2.StringValue);

        var row3 = sheet.ReadRow<DictionaryClass>();
        Assert.Equal("B", row3.StringValue);

        var row4 = sheet.ReadRow<DictionaryClass>();
        Assert.Null(row4.StringValue);
    }

    [Fact]
    public void ReadRow_CustomMappedMultipleMappers_Success()
    {
        using var importer = Helpers.GetImporter("WithMappings.xlsx");
        importer.Configuration.RegisterClassMap<DictionaryClass>(c =>
        {
            c.Map(m => m.StringValue)
                .WithMappers(new StringMapper());
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<DictionaryClass>();
        Assert.Equal("a", row1.StringValue);

        var row2 = sheet.ReadRow<DictionaryClass>();
        Assert.Equal("b", row2.StringValue);

        var row3 = sheet.ReadRow<DictionaryClass>();
        Assert.Equal("B", row3.StringValue);

        var row4 = sheet.ReadRow<DictionaryClass>();
        Assert.Null(row4.StringValue);
    }

    private class DictionaryClass
    {
        [ExcelMapper(typeof(StringMapper))]
        [ExcelMapper(typeof(CustomMapper), ConstructorArguments = [true])]
        public string StringValue { get; set; } = default!;
    }

    private class CustomMapper : ICellMapper
    {
        private ICellMapper _innerMapper;

        public CustomMapper(bool value)
        {
            Assert.True(value);
            _innerMapper = new MappingDictionaryMapper<string>(new Dictionary<string, string>
                    {
                        { "a", "12345" }
                    }, null, behavior: MappingDictionaryMapperBehavior.Optional);
        }

        public CellMapperResult Map(ReadCellResult readResult) => _innerMapper.Map(readResult);
    }
}
