namespace ExcelMapper.Tests;

public class MapUInt16Tests
{
    [Fact]
    public void ReadRow_UInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ushort>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort>());
    }

    [Fact]
    public void ReadRow_NullableUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ushort?>();
        Assert.Equal((ushort)2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<ushort?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt16Class>();
        Assert.Equal((ushort)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt16Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultUInt16ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableUInt16ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt16Class>();
        Assert.Equal((ushort)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt16Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt16Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomUInt16ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt16Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt16Value>();
        Assert.Equal(11, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<UInt16Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt16_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableUInt16ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt16Class>();
        Assert.Equal((ushort)2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt16Class>();
        Assert.Equal((ushort)11, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableUInt16Class>();
        Assert.Equal((ushort)10, row3.Value);
    }

    [Fact]
    public void ReadRow_UInt16Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ushort>());
    }

    private class UInt16Value
    {
        public ushort Value { get; set; }
    }

    private class DefaultUInt16ValueMap : ExcelClassMap<UInt16Value>
    {
        public DefaultUInt16ValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomUInt16ValueMap : ExcelClassMap<UInt16Value>
    {
        public CustomUInt16ValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }

    private class NullableUInt16Class
    {
        public ushort? Value { get; set; }
    }

    private class DefaultNullableUInt16ClassMap : ExcelClassMap<NullableUInt16Class>
    {
        public DefaultNullableUInt16ClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableUInt16ClassMap : ExcelClassMap<NullableUInt16Class>
    {
        public CustomNullableUInt16ClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }
}
