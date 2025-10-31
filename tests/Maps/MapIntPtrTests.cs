using System.Globalization;

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

    [Fact]
    public void ReadRow_AutoMappedIntPtrValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeIntPtrValue>();
        Assert.Equal((nint)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeIntPtrValue>();
        Assert.Equal((nint)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeIntPtrValue>();
        Assert.Equal((nint)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeIntPtrValue>();
        Assert.Equal((nint)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeIntPtrValue>();
        Assert.Equal((nint)123, row5.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeIntPtrValue>());

        // Invalid cell values.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeIntPtrValue>());
    }

    private class NumberStyleAttributeIntPtrValue
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        public nint Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedIntPtrValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeIntPtrValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeIntPtrValue>();
        Assert.Equal((nint)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeIntPtrValue>();
        Assert.Equal((nint)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeIntPtrValue>();
        Assert.Equal((nint)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeIntPtrValue>();
        Assert.Equal((nint)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeIntPtrValue>();
        Assert.Equal((nint)123, row5.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeIntPtrValue>());

        // Invalid cell values.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeIntPtrValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableIntPtrValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Equal((nint)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Equal((nint)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Equal((nint)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Equal((nint)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Equal((nint)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Null(row6.Value);

        // Invalid cell values.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>());
    }

    private class NumberStyleAttributeNullableIntPtrValue
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        public nint? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableIntPtrValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableIntPtrValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Equal((nint)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Equal((nint)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Equal((nint)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Equal((nint)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Equal((nint)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>();
        Assert.Null(row6.Value);

        // Invalid cell values.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableIntPtrValue>());
    }
}
