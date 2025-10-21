using System;
using System.Diagnostics.CodeAnalysis;
using Xunit;

namespace ExcelMapper.Tests;

public class MapIParsableTests
{
    [Fact]
    public void ReadRow_IParsableClass_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableClass>();
        Assert.Equal("value", row1.ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableClass>();
        Assert.Equal("  value  ", row2.ParsedMember);

        // Empty value
        var row3 = sheet.ReadRow<ParsableClass>();
        Assert.Null(row3);
    }

    private class ParsableClass : IParsable<ParsableClass>
    {
        public string? ParsedMember { get; set; }

        public static ParsableClass Parse(string s, IFormatProvider? provider)
        {
            return new ParsableClass { ParsedMember = s };
        }

        public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, [MaybeNullWhen(false)] out ParsableClass result)
            => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_IParsableStruct_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableClass>();
        Assert.Equal("value", row1.ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableClass>();
        Assert.Equal("  value  ", row2.ParsedMember);

        // Empty value
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ParsableStruct>());
    }

    private struct ParsableStruct : IParsable<ParsableStruct>
    {
        public string? ParsedMember { get; set; }

        public static ParsableStruct Parse(string s, IFormatProvider? provider)
        {
            return new ParsableStruct { ParsedMember = s };
        }

        public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, [MaybeNullWhen(false)] out ParsableStruct result)
            => throw new NotImplementedException();
    }

    [Fact]
    public void ReadRow_IParsableNullableStruct_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableClass?>();
        Assert.Equal("value", row1!.ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableClass?>();
        Assert.Equal("  value  ", row2!.ParsedMember);

        // Empty value
        var row3 = sheet.ReadRow<ParsableClass?>();
        Assert.Null(row3);
    }

    [Fact]
    public void ReadRow_AutoMappedIParsableClassValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableClassValue>();
        Assert.Equal("value", Assert.IsType<ParsableClass>(row1.Value).ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableClassValue>();
        Assert.Equal("  value  ", Assert.IsType<ParsableClass>(row2.Value).ParsedMember);

        // Empty value
        var row3 = sheet.ReadRow<ParsableClassValue>();
        Assert.Null(row3.Value);
    }

    private class ParsableClassValue
    {
        public ParsableClass Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedIParsableClassValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ParsableClassValue>(c =>
        {
            c.Map(o => o.Value);
        });

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableClassValue>();
        Assert.Equal("value", row1.Value.ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableClassValue>();
        Assert.Equal("  value  ", row2.Value.ParsedMember);

        // Empty value
        var row3 = sheet.ReadRow<ParsableClassValue>();
        Assert.Null(row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIParsableClassValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ParsableClassValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new ParsableClass { ParsedMember = "empty" })
                .WithInvalidFallback(new ParsableClass { ParsedMember = "invalid" });
        });

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableClassValue>();
        Assert.Equal("value", row1.Value.ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableClassValue>();
        Assert.Equal("  value  ", row2.Value.ParsedMember);

        // Empty value
        var row3 = sheet.ReadRow<ParsableClassValue>();
        Assert.Equal("empty", row3.Value.ParsedMember);
    }

    [Fact]
    public void ReadRow_AutoMappedIParsableStructValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableStructValue>();
        Assert.Equal("value", Assert.IsType<ParsableStruct>(row1.Value).ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableStructValue>();
        Assert.Equal("  value  ", Assert.IsType<ParsableStruct>(row2.Value).ParsedMember);

        // Empty value
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ParsableStructValue>());
    }

    private struct ParsableStructValue
    {
        public ParsableStructValue()
        {
        }

        public ParsableStruct Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedIParsableStructValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ParsableStructValue>(c =>
        {
            c.Map(o => o.Value);
        });

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableStructValue>();
        Assert.Equal("value", row1.Value.ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableStructValue>();
        Assert.Equal("  value  ", row2.Value.ParsedMember);

        // Empty value
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ParsableStructValue>());
    }

    [Fact]
    public void ReadRow_CustomMappedIParsableStructValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ParsableStructValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new ParsableStruct { ParsedMember = "empty" })
                .WithInvalidFallback(new ParsableStruct { ParsedMember = "invalid" });
        });

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableStructValue>();
        Assert.Equal("value", row1.Value.ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableStructValue>();
        Assert.Equal("  value  ", row2.Value.ParsedMember);

        // Empty value
        var row3 = sheet.ReadRow<ParsableStructValue>();
        Assert.Equal("empty", row3.Value.ParsedMember);
    }

    [Fact]
    public void ReadRow_AutoMappedIParsableNullableStructValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableNullableStructValue>();
        Assert.Equal("value", Assert.IsType<ParsableStruct>(row1.Value).ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableNullableStructValue>();
        Assert.Equal("  value  ", Assert.IsType<ParsableStruct>(row2.Value).ParsedMember);

        // Empty value
        var row3 = sheet.ReadRow<ParsableNullableStructValue>();
        Assert.Null(row3.Value);
    }

    private struct ParsableNullableStructValue
    {
        public ParsableNullableStructValue()
        {
        }

        public ParsableStruct? Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedIParsableNullableStructValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ParsableNullableStructValue>(c =>
        {
            c.Map(o => o.Value);
        });

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableNullableStructValue>();
        Assert.Equal("value", row1.Value!.Value.ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableNullableStructValue>();
        Assert.Equal("  value  ", row2.Value!.Value.ParsedMember);

        // Empty value
        var row3 = sheet.ReadRow<ParsableNullableStructValue>();
        Assert.Null(row3.Value);
    }

    [Fact]
    public void ReadRow_CustomMappedIParsableNullableStructValue_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<ParsableNullableStructValue>(c =>
        {
            c.Map(o => o.Value)
                .WithEmptyFallback(new ParsableStruct { ParsedMember = "empty" })
                .WithInvalidFallback(new ParsableStruct { ParsedMember = "invalid" });
        });

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid value
        var row1 = sheet.ReadRow<ParsableNullableStructValue>();
        Assert.Equal("value", row1.Value!.Value.ParsedMember);

        // Valid value
        var row2 = sheet.ReadRow<ParsableNullableStructValue>();
        Assert.Equal("  value  ", row2.Value!.Value.ParsedMember);

        // Empty value
        var row3 = sheet.ReadRow<ParsableNullableStructValue>();
        Assert.Equal("empty", row3.Value!.Value.ParsedMember);
    }
}
