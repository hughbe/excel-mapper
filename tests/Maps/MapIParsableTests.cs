// These types are only supported by the library's .NET 7+ build.
#if LIBRARY_NET8
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ExcelMapper.Tests;

public class MapIParsableTests
{
    [Fact]
    public void ReadRow_IParsableClass_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
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
    public void ReadRows_IParsableClassWithoutParameterlessConstructor_Parses()
    {
        // Created by its Parse method, so it needn't have a parameterless constructor: it isn't
        // mapped through its members.
        using var importer = Helpers.GetCsvImporter("Code,Value\nab,x\n");
        var row = importer.ReadSheet().ReadRows<CodeRow>().Single();
        Assert.Equal("AB", row.Code.Value);
        Assert.Equal("AB", row.Abstract!.Value);
    }

    [Fact]
    public void ReadRows_IPAddress_Parses()
    {
        using var importer = Helpers.GetCsvImporter("Address\n192.168.0.1\n");
        Assert.Equal(System.Net.IPAddress.Parse("192.168.0.1"), importer.ReadSheet().ReadRows<AddressRow>().Single().Address);
    }

    private class CodeRow
    {
        public Code Code { get; set; } = new("");

        [ExcelColumnName("Code")]
        public AbstractCode? Abstract { get; set; }
    }

    private class AddressRow
    {
        public System.Net.IPAddress? Address { get; set; }
    }

    private sealed record Code(string Value) : IParsable<Code>
    {
        public static Code Parse(string s, IFormatProvider? provider) => new(s.ToUpperInvariant());

        public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, [MaybeNullWhen(false)] out Code result)
        {
            result = s is null ? null : Parse(s, provider);
            return result is not null;
        }
    }

    // An abstract type is created by its Parse method too.
    private abstract class AbstractCode : IParsable<AbstractCode>
    {
        public abstract string Value { get; }

        public static AbstractCode Parse(string s, IFormatProvider? provider) => new UpperCode(s.ToUpperInvariant());

        public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, [MaybeNullWhen(false)] out AbstractCode result)
        {
            result = s is null ? null : Parse(s, provider);
            return result is not null;
        }
    }

    private sealed class UpperCode(string value) : AbstractCode
    {
        public override string Value { get; } = value;
    }

    [Fact]
    public void ReadRow_IParsableStruct_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");

        var sheet = importer.ReadSheet();
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

        var sheet = importer.ReadSheet();
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

        var sheet = importer.ReadSheet();
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

        var sheet = importer.ReadSheet();
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

        var sheet = importer.ReadSheet();
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

        var sheet = importer.ReadSheet();
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

        var sheet = importer.ReadSheet();
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

        var sheet = importer.ReadSheet();
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

        var sheet = importer.ReadSheet();
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

        var sheet = importer.ReadSheet();
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

        var sheet = importer.ReadSheet();
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
#endif
