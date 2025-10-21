using Xunit;

namespace ExcelMapper.Tests;

public class MapIntPtrTests
{
    [Fact]
    public void ReadRow_IntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<nint>();
        Assert.Equal(2, row1);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<nint>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<nint>());
    }

    [Fact]
    public void ReadRow_NullableIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<nint?>();
        Assert.Equal(2, row1);

        // Empty cell value.
        var row2 = sheet.ReadRow<nint?>();
        Assert.Null(row2);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntPtrValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<IntPtrValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntPtrValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntPtrValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableIntPtrClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableIntPtrClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntPtrValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultIntPtrValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<IntPtrValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntPtrValue>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntPtrValue>());
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<DefaultNullableIntPtrClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableIntPtrClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableIntPtrClass>();
        Assert.Null(row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntPtrValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomIntPtrValueMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<IntPtrValue>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<IntPtrValue>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<IntPtrValue>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedNullableIntPtr_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<CustomNullableIntPtrClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<NullableIntPtrClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<NullableIntPtrClass>();
        Assert.Equal(-10, row2.Value);

        // Invalid cell value.
        var row3 = sheet.ReadRow<NullableIntPtrClass>();
        Assert.Equal(10, row3.Value);
    }

    [Fact]
    public void ReadRow_IntPtrOverflow_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Overflow_Signed.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<nint>());
    }

    private class IntPtrValue
    {
        public nint Value { get; set; }
    }

    private class DefaultIntPtrValueMap : ExcelClassMap<IntPtrValue>
    {
        public DefaultIntPtrValueMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomIntPtrValueMap : ExcelClassMap<IntPtrValue>
    {
        public CustomIntPtrValueMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        }
    }

    private class NullableIntPtrClass
    {
        public nint? Value { get; set; }
    }

    private class DefaultNullableIntPtrClassMap : ExcelClassMap<NullableIntPtrClass>
    {
        public DefaultNullableIntPtrClassMap()
        {
            Map(o => o.Value);
        }
    }

    private class CustomNullableIntPtrClassMap : ExcelClassMap<NullableIntPtrClass>
    {
        public CustomNullableIntPtrClassMap()
        {
            Map(o => o.Value)
                .WithEmptyFallback(-10)
                .WithInvalidFallback(10);
        }
    }
}
