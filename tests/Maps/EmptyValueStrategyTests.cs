using System;
using Xunit;

namespace ExcelMapper.Tests;

public class EmptyValueStrategyTests
{
    [Fact]
    public void ReadRow_AutoMappedStringEmptyValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("fallback", row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedStringEmptyNullValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringNullClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedStringEmptyInvalidValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        Assert.Throws<InvalidCastException>(() => sheet.ReadRow<EmptyValueStringInvalidClass>());

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringInvalidClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedStringEmptyValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueStringClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("fallback", row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedStringEmptyNullValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueStringNullClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        var row3 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Null(row3.Value);

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringNullClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringNullClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedStringEmptyInvalidValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueStringInvalidClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("value", row1.Value);

        // Valid value
        var row2 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("  value  ", row2.Value);

        // Empty value
        Assert.Throws<InvalidCastException>(() => sheet.ReadRow<EmptyValueStringInvalidClass>());

        // Last row.
        var row4 = sheet.ReadRow<EmptyValueStringInvalidClass>();
        Assert.Equal("value", row4.Value);

        // No more rows.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueStringInvalidClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedIntEmptyValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<EmptyValueIntClass>();
        Assert.Equal(1, row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedIntEmptyNullValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntNullClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<NullReferenceException>(() => sheet.ReadRow<EmptyValueIntNullClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntNullClass>());
    }

    [Fact]
    public void ReadRow_AutoMappedEmptyInvalidValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntInvalidClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<InvalidCastException>(() => sheet.ReadRow<EmptyValueIntInvalidClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntInvalidClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedEmptyValueInt_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueIntClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        var row2 = sheet.ReadRow<EmptyValueIntClass>();
        Assert.Equal(1, row2.Value);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedEmptyNullValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueIntNullClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntNullClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<NullReferenceException>(() => sheet.ReadRow<EmptyValueIntNullClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntNullClass>());
    }

    [Fact]
    public void ReadRow_DefaultMappedEmptyInvalidValue_Success()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap<EmptyValueIntInvalidClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value.
        var row1 = sheet.ReadRow<EmptyValueIntInvalidClass>();
        Assert.Equal(2, row1.Value);

        // Empty cell value.
        Assert.Throws<InvalidCastException>(() => sheet.ReadRow<EmptyValueIntInvalidClass>());

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EmptyValueIntInvalidClass>());
    }

    [Fact]
    public void ReadRow_EmptyValueStrategy_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("EmptyValues.xlsx");
        importer.Configuration.RegisterClassMap(new EmptyValueStrategyMap());

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<EmptyValues>();
        Assert.Equal(0, row1.IntValue);
        Assert.Null(row1.StringValue);
        Assert.False(row1.BoolValue);
        Assert.Equal((EmptyValuesEnum)0, row1.EnumValue);
        Assert.Equal(DateTime.MinValue, row1.DateValue);
        Assert.Equal([0, 0], row1.ArrayValue);
    }

    public class EmptyValueStringClass
    {
        [ExcelDefaultValue("fallback")]
        public string? Value { get; set; }
    }

    public class EmptyValueStringNullClass
    {
        [ExcelDefaultValue(null)]
        public string? Value { get; set; }
    }

    public class EmptyValueStringInvalidClass
    {
        [ExcelDefaultValue(1)]
        public string? Value { get; set; }
    }

    public class EmptyValueStringClassMap : ExcelClassMap<EmptyValueStringClass>
    {
        public EmptyValueStringClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class EmptyValueStringNullClassMap : ExcelClassMap<EmptyValueStringNullClass>
    {
        public EmptyValueStringNullClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class EmptyValueStringInvalidClassMap : ExcelClassMap<EmptyValueStringInvalidClass>
    {
        public EmptyValueStringInvalidClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class EmptyValueIntClass
    {
        [ExcelDefaultValue(1)]
        public int Value { get; set; }
    }

    public class EmptyValueIntNullClass
    {
        [ExcelDefaultValue(null)]
        public int Value { get; set; }
    }

    public class EmptyValueIntInvalidClass
    {
        [ExcelDefaultValue("fallback")]
        public int Value { get; set; }
    }

    public class EmptyValueIntClassMap : ExcelClassMap<EmptyValueIntClass>
    {
        public EmptyValueIntClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class EmptyValueIntNullClassMap : ExcelClassMap<EmptyValueIntNullClass>
    {
        public EmptyValueIntNullClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class EmptyValueIntInvalidClassMap : ExcelClassMap<EmptyValueIntInvalidClass>
    {
        public EmptyValueIntInvalidClassMap()
        {
            Map(o => o.Value);
        }
    }

    public class EmptyValues
    {
        public int IntValue { get; set; }
        public string StringValue { get; set; } = default!;
        public bool BoolValue { get; set; }
        public EmptyValuesEnum EnumValue { get; set; }
        public DateTime DateValue { get; set; }
        public int[] ArrayValue { get; set; } = default!;
    }

    public enum EmptyValuesEnum
    {
        Test = 1
    }

    public class EmptyValueStrategyMap : ExcelClassMap<EmptyValues>
    {
        public EmptyValueStrategyMap() : base(FallbackStrategy.SetToDefaultValue)
        {
            Map(e => e.IntValue);
            Map(e => e.StringValue);
            Map(e => e.BoolValue);
            Map(e => e.EnumValue);
            Map(e => e.DateValue);
            Map(e => e.ArrayValue)
                .WithColumnNames("ArrayValue1", "ArrayValue2");
        }
    }
}
