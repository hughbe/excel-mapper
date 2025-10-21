using Xunit;

namespace ExcelMapper.Tests;

public class MapUIntPtrTests
{
    [Fact]
    public void ReadRow_UIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<nuint>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<nuint>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<nuint>());
    }

    [Fact]
    public void ReadRow_NullableUIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<nuint?>();
        Assert.Equal(2u, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<nuint?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UIntPtrValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedUIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UIntPtrValue>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UIntPtrValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UIntPtrValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUIntPtrClass>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUIntPtrClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UIntPtrValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedUIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultUIntPtrValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UIntPtrValue>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UIntPtrValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UIntPtrValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableUIntPtrClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUIntPtrClass>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUIntPtrClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UIntPtrValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedUIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomUIntPtrValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<UIntPtrValue>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<UIntPtrValue>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<UIntPtrValue>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableUIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableUIntPtrClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableUIntPtrClass>();
        Assert.Equal(2u, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableUIntPtrClass>();
        Assert.Equal(11u, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableUIntPtrClass>();
        Assert.Equal(10u, row3.Value);
    }

    [Fact]
    public void ReadRow_UIntPtrOverflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<nuint>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<nuint>());
    }

    private class UIntPtrValue
    {
        public nuint Value { get; set; }
    }

    private class DefaultUIntPtrValueMap : ExcelClassMap<UIntPtrValue>
    {
        public DefaultUIntPtrValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomUIntPtrValueMap : ExcelClassMap<UIntPtrValue>
    {
        public CustomUIntPtrValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }

    private class NullableUIntPtrClass
    {
        public nuint? Value { get; set; }
    }

    private class DefaultNullableUIntPtrClassMap : ExcelClassMap<NullableUIntPtrClass>
    {
        public DefaultNullableUIntPtrClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableUIntPtrClassMap : ExcelClassMap<NullableUIntPtrClass>
    {
        public CustomNullableUIntPtrClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(11)
                .WithInvalidFallback(10);
        }
    }
}
