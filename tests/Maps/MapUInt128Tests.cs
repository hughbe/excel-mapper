using System;
using Xunit;

namespace ExcelMapper.Tests;

public class MapUInt128Tests
{
    [Fact]
    public void ReadRow_UInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt128>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128>());
    }

    [Fact]
    public void ReadRow_NullableUInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt128?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt128?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedUInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt128Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt128Class>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt128Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultUInt128ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt128Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableUInt128ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt128Class>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt128Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedUInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomUInt128ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UInt128Value>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UInt128Value>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<UInt128Value>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableUInt128ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUInt128Class>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUInt128Class>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableUInt128Class>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_UInt128Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UInt128>());
    }

    private class UInt128Value
    {
        public UInt128 Value { get; set; }
    }

    private class DefaultUInt128ValueMap : ExcelClassMap<UInt128Value>
    {
        public DefaultUInt128ValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomUInt128ValueMap : ExcelClassMap<UInt128Value>
    {
        public CustomUInt128ValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }

    private class NullableUInt128Class
    {
        public UInt128? Value { get; set; }
    }

    private class DefaultNullableUInt128ClassMap : ExcelClassMap<NullableUInt128Class>
    {
        public DefaultNullableUInt128ClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableUInt128ClassMap : ExcelClassMap<NullableUInt128Class>
    {
        public CustomNullableUInt128ClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback((UInt128)11)
                .WithInvalidFallback((UInt128)10);
        }
    }
}
