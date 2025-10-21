using Xunit;

namespace ExcelMapper.Tests;

public class MapUInt64Tests
{
    [Fact]
    public void ReadRow_UInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ulong>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong>());
    }

    [Fact]
    public void ReadRow_NullableUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<ulong?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<ulong?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt64Class>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt64Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultUInt64ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableUInt64ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt64Class>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt64Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt64Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomUInt64ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<UInt64Value>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt64_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableUInt64ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt64Class>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt64Class>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableUInt64Class>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_UInt64Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ulong>());
    }

    private class UInt64Value
    {
        public ulong Value { get; set; }
    }

    private class DefaultUInt64ValueMap : ExcelClassMap<UInt64Value>
    {
        public DefaultUInt64ValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomUInt64ValueMap : ExcelClassMap<UInt64Value>
    {
        public CustomUInt64ValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }

    private class NullableUInt64Class
    {
        public ulong? Value { get; set; }
    }

    private class DefaultNullableUInt64ClassMap : ExcelClassMap<NullableUInt64Class>
    {
        public DefaultNullableUInt64ClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableUInt64ClassMap : ExcelClassMap<NullableUInt64Class>
    {
        public CustomNullableUInt64ClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback((ulong)11)
                .WithInvalidFallback((ulong)10);
        }
    }
}
