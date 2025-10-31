using System.Globalization;

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

    [Fact]
    public void ReadRow_AutoMappedUIntPtrValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeUIntPtrValue>();
        Assert.Equal((nuint)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeUIntPtrValue>();
        Assert.Equal((nuint)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeUIntPtrValue>();
        Assert.Equal((nuint)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeUIntPtrValue>();
        Assert.Equal((nuint)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeUIntPtrValue>();
        Assert.Equal((nuint)123, row5.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeUIntPtrValue>());

        // Invalid cell values.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeUIntPtrValue>());
        
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeUIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeUIntPtrValue>());
    }

    private class NumberStyleAttributeUIntPtrValue
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        public nuint Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedUIntPtrValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeUIntPtrValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeUIntPtrValue>();
        Assert.Equal((nuint)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeUIntPtrValue>();
        Assert.Equal((nuint)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeUIntPtrValue>();
        Assert.Equal((nuint)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeUIntPtrValue>();
        Assert.Equal((nuint)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeUIntPtrValue>();
        Assert.Equal((nuint)123, row5.Value);

        // Empty cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeUIntPtrValue>());

        // Invalid cell values.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeUIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeUIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeUIntPtrValue>());
    }

    [Fact]
    public void ReadRow_AutoMappedNullableUIntPtrValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Equal((nuint)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Equal((nuint)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Equal((nuint)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Equal((nuint)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Equal((nuint)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Null(row6.Value);

        // Invalid cell values.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>());
    }

    private class NumberStyleAttributeNullableUIntPtrValue
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        public nuint? Value { get; set; }
    }

    [Fact]
    public void ReadRow_DefaultMappedNullableUIntPtrValueNumberStyleAttribute_Success()
    {
        using var importer = Helpers.GetImporter("Numbers_Hex.xlsx");
        importer.Configuration.RegisterClassMap<NumberStyleAttributeNullableUIntPtrValue>(c =>
        {
            c.Map(o => o.Value);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell values.
        var row1 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Equal((nuint)0xAB, row1.Value);

        var row2 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Equal((nuint)0x123, row2.Value);

        var row3 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Equal((nuint)0xAB, row3.Value);

        var row4 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Equal((nuint)0x123, row4.Value);

        var row5 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Equal((nuint)123, row5.Value);

        // Empty cell value.
        var row6 = sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>();
        Assert.Null(row6.Value);

        // Invalid cell values.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>());

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NumberStyleAttributeNullableUIntPtrValue>());
    }
}
