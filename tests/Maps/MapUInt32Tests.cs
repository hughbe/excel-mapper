using Xunit;

namespace ExcelMapper.Tests;

public class MapUInt32Tests
{
    [Fact]
    public void ReadRow_UInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<uint>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint>());
    }

    [Fact]
    public void ReadRow_NullableUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<uint?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<uint?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt32Class>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt32Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultUInt32ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableUInt32ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt32Class>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt32Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt32Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomUInt32ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<UInt32Value>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt32_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableUInt32ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt32Class>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt32Class>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableUInt32Class>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_UInt32Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<uint>());
    }

    private class UInt32Value
    {
        public uint Value { get; set; }
    }

    private class DefaultUInt32ValueMap : ExcelClassMap<UInt32Value>
    {
        public DefaultUInt32ValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomUInt32ValueMap : ExcelClassMap<UInt32Value>
    {
        public CustomUInt32ValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }

    private class NullableUInt32Class
    {
        public uint? Value { get; set; }
    }

    private class DefaultNullableUInt32ClassMap : ExcelClassMap<NullableUInt32Class>
    {
        public DefaultNullableUInt32ClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableUInt32ClassMap : ExcelClassMap<NullableUInt32Class>
    {
        public CustomNullableUInt32ClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }
}
