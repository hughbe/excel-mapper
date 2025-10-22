namespace ExcelMapper.Tests;

public class MapBoolTests
{
    [Fact]
    public void ReadRow_Bool_Success()
    {
        using var importer = Helpers.GetImporter("Bools.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<bool>();
        Assert.True(row1);

        var row2 = sheet.ReadRow<bool>();
        Assert.True(row2);

        var row3 = sheet.ReadRow<bool>();
        Assert.False(row3);

        var row4 = sheet.ReadRow<bool>();
        Assert.False(row4);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<bool>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<bool>());
    }

    [Fact]
    public void ReadRow_NullableBool_Success()
    {
        using var importer = Helpers.GetImporter("Bools.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<bool?>();
        Assert.True(row1);

        var row2 = sheet.ReadRow<bool?>();
        Assert.True(row2);

        var row3 = sheet.ReadRow<bool?>();
        Assert.False(row3);

        var row4 = sheet.ReadRow<bool?>();
        Assert.False(row4);

        // Empty cell value.
        var row5 = sheet.ReadRow<bool?>();
        Assert.Null(row5);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<bool?>());
    }

    [Fact]
    public void ReadRow_AutoMappedBool_Success()
    {
        using var importer = Helpers.GetImporter("Bools.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<BoolValue>();
        Assert.True(row1.Value);

        var row2 = sheet.ReadRow<BoolValue>();
        Assert.True(row2.Value);

        var row3 = sheet.ReadRow<BoolValue>();
        Assert.False(row3.Value);

        var row4 = sheet.ReadRow<BoolValue>();
        Assert.False(row4.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BoolValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BoolValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableBool_Success()
    {
        using var importer = Helpers.GetImporter("Bools.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableBoolValue>();
        Assert.True(row1.Value);

        var row2 = sheet.ReadRow<NullableBoolValue>();
        Assert.True(row2.Value);

        var row3 = sheet.ReadRow<NullableBoolValue>();
        Assert.False(row3.Value);

        var row4 = sheet.ReadRow<NullableBoolValue>();
        Assert.False(row4.Value);

        // Empty cell value.
        var row5 = sheet.ReadRow<NullableBoolValue>();
        Assert.Null(row5.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableBoolValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedBool_Success()
    {
        using var importer = Helpers.GetImporter("Bools.xlsx");
        importer.Configuration.RegisterClassMap<DefaultBoolClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<BoolValue>();
        Assert.True(row1.Value);

        var row2 = sheet.ReadRow<BoolValue>();
        Assert.True(row2.Value);

        var row3 = sheet.ReadRow<BoolValue>();
        Assert.False(row3.Value);

        var row4 = sheet.ReadRow<BoolValue>();
        Assert.False(row4.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BoolValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BoolValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableBool_Success()
    {
        using var importer = Helpers.GetImporter("Bools.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableBoolMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableBoolValue>();
        Assert.True(row1.Value);

        var row2 = sheet.ReadRow<NullableBoolValue>();
        Assert.True(row2.Value);

        var row3 = sheet.ReadRow<NullableBoolValue>();
        Assert.False(row3.Value);

        var row4 = sheet.ReadRow<NullableBoolValue>();
        Assert.False(row4.Value);

        // Empty cell value.
        var row5 = sheet.ReadRow<NullableBoolValue>();
        Assert.Null(row5.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableBoolValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedBool_Success()
    {
        using var importer = Helpers.GetImporter("Bools.xlsx");
        importer.Configuration.RegisterClassMap<CustomBoolClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<BoolValue>();
        Assert.True(row1.Value);

        var row2 = sheet.ReadRow<BoolValue>();
        Assert.True(row2.Value);

        var row3 = sheet.ReadRow<BoolValue>();
        Assert.False(row3.Value);

        var row4 = sheet.ReadRow<BoolValue>();
        Assert.False(row4.Value);

        // Empty cell value.
        var row5 = sheet.ReadRow<BoolValue>();
        Assert.True(row5.Value);

        // Invalid cell value.
        var row6 = sheet.ReadRow<BoolValue>();
        Assert.True(row6.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableBool_Success()
    {
        using var importer = Helpers.GetImporter("Bools.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableBoolMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableBoolValue>();
        Assert.True(row1.Value);

        var row2 = sheet.ReadRow<NullableBoolValue>();
        Assert.True(row2.Value);

        var row3 = sheet.ReadRow<NullableBoolValue>();
        Assert.False(row3.Value);

        var row4 = sheet.ReadRow<NullableBoolValue>();
        Assert.False(row4.Value);

        // Empty cell value.
        var row5 = sheet.ReadRow<NullableBoolValue>();
        Assert.True(row5.Value);

        // Invalid cell value.
        var row6 = sheet.ReadRow<NullableBoolValue>();
        Assert.True(row6.Value);
    }

    private class BoolValue
    {
        public bool Value { get; set; }
    }

    private class DefaultBoolClassMap : ExcelClassMap<BoolValue>
    {
        public DefaultBoolClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomBoolClassMap : ExcelClassMap<BoolValue>
    {
        public CustomBoolClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(true)
                .WithInvalidFallback(true);
        }
    }

    private class NullableBoolValue
    {
        public bool? Value { get; set; }
    }

    private class DefaultNullableBoolMap : ExcelClassMap<NullableBoolValue>
    {
        public DefaultNullableBoolMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableBoolMap : ExcelClassMap<NullableBoolValue>
    {
        public CustomNullableBoolMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(true)
                .WithInvalidFallback(true);
        }
    }
}
