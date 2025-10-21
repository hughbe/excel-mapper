using System;
using Xunit;

namespace ExcelMapper.Tests;

public class MapInt128Tests
{
    [Fact]
    public void ReadRow_Int128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int128>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int128>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int128>());
    }

    [Fact]
    public void ReadRow_NullableInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int128?>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<Int128?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int128Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int128Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int128Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int128Value>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt128Class>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt128Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int128Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultInt128ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int128Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int128Value>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int128Value>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableInt128ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt128Class>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt128Class>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int128Value>());
    }

    [Fact]
    public void ReadRow_CustomMappedInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomInt128ValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<Int128Value>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<Int128Value>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<Int128Value>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableInt128_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableInt128ClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableInt128Class>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableInt128Class>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableInt128Class>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_Int128Overflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int128>());
    }

    private class Int128Value
    {
        public Int128 Value { get; set; }
    }

    private class DefaultInt128ValueMap : ExcelClassMap<Int128Value>
    {
        public DefaultInt128ValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomInt128ValueMap : ExcelClassMap<Int128Value>
    {
        public CustomInt128ValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        }
    }

    private class NullableInt128Class
    {
        public Int128? Value { get; set; }
    }

    private class DefaultNullableInt128ClassMap : ExcelClassMap<NullableInt128Class>
    {
        public DefaultNullableInt128ClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableInt128ClassMap : ExcelClassMap<NullableInt128Class>
    {
        public CustomNullableInt128ClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        }
    }
}
