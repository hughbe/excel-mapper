using System.Globalization;
using System.Linq;
using System.Numerics;

namespace ExcelMapper.Tests;

/// <summary>
/// Formats, format providers, number styles, URI kinds and offsets belong to each member's map, and
/// are given to the built-in conversion with each cell. The built-in mappers are shared by every
/// member of a type, so these check that one member's options don't affect another's.
/// </summary>
public class MapConversionOptionsTests
{
    public class TwoDates
    {
        public DateTime First { get; set; }
        public DateTime Second { get; set; }
        public DateTime Third { get; set; }
    }

    [Fact]
    public void ReadRows_MembersOfOneTypeWithDifferentFormats_UseTheirOwn()
    {
        using var importer = Helpers.GetCsvImporter("First,Second,Third\n20240304,04/03/2024,2024-03-04\n");
        importer.Configuration.RegisterClassMap<TwoDates>(c =>
        {
            c.Map(r => r.First).WithFormats("yyyyMMdd");
            c.Map(r => r.Second).WithFormats("dd/MM/yyyy");
            c.Map(r => r.Third);
        });

        var row = importer.ReadSheet().ReadRow<TwoDates>();
        Assert.Equal(new DateTime(2024, 3, 4), row.First);
        Assert.Equal(new DateTime(2024, 3, 4), row.Second);
        Assert.Equal(new DateTime(2024, 3, 4), row.Third);
    }

    [Fact]
    public void ReadRows_FormatsOfOneMember_DontApplyToAnother()
    {
        // Only First has the format, so Second can't be read with it.
        using var importer = Helpers.GetCsvImporter("First,Second,Third\n20240304,20240304,2024-03-04\n");
        importer.Configuration.RegisterClassMap<TwoDates>(c =>
        {
            c.Map(r => r.First).WithFormats("yyyyMMdd");
            c.Map(r => r.Second).WithInvalidFallback(DateTime.MinValue);
            c.Map(r => r.Third);
        });

        var row = importer.ReadSheet().ReadRow<TwoDates>();
        Assert.Equal(new DateTime(2024, 3, 4), row.First);
        Assert.Equal(DateTime.MinValue, row.Second);
    }

    public class TwoDatesWithAttributes
    {
        [ExcelFormats("yyyyMMdd")]
        public DateTime First { get; set; }

        public DateTime Second { get; set; }
    }

    [Fact]
    public void ReadRows_AttributeFormatsOnOneMember_DontApplyToAnother()
    {
        using var importer = Helpers.GetCsvImporter("First,Second\n20240304,2024-03-04\n20240305,20240305\n");
        var sheet = importer.ReadSheet();

        var row = sheet.ReadRow<TwoDatesWithAttributes>();
        Assert.Equal(new DateTime(2024, 3, 4), row.First);
        Assert.Equal(new DateTime(2024, 3, 4), row.Second);

        // "20240305" isn't one of Second's formats.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<TwoDatesWithAttributes>());
    }

    public class DateLists
    {
        public DateTime[] First { get; set; } = default!;
        public DateTime[] Second { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_ElementsWithDifferentFormats_UseTheirOwn()
    {
        using var importer = Helpers.GetCsvImporter("First,Second\n\"20240304,20240305\",\"04/03/2024,05/03/2024\"\n");
        importer.Configuration.RegisterClassMap<DateLists>(c =>
        {
            c.Map(r => r.First).WithElementMap(e => e.WithFormats("yyyyMMdd"));
            c.Map(r => r.Second).WithElementMap(e => e.WithFormats("dd/MM/yyyy"));
        });

        var row = importer.ReadSheet().ReadRow<DateLists>();
        Assert.Equal([new DateTime(2024, 3, 4), new DateTime(2024, 3, 5)], row.First);
        Assert.Equal([new DateTime(2024, 3, 4), new DateTime(2024, 3, 5)], row.Second);
    }

    public class TwoNumbers
    {
        public int First { get; set; }
        public int Second { get; set; }
    }

    [Fact]
    public void ReadRows_MembersOfOneTypeWithDifferentProviders_UseTheirOwn()
    {
        var german = CultureInfo.GetCultureInfo("de-DE");
        using var importer = Helpers.GetCsvImporter("First,Second\n\"1.234\",\"1,234\"\n");
        importer.Configuration.Culture = CultureInfo.InvariantCulture;
        importer.Configuration.RegisterClassMap<TwoNumbers>(c =>
        {
            c.Map(r => r.First).WithFormatProvider(german);
            c.Map(r => r.Second);
        });

        var row = importer.ReadSheet().ReadRow<TwoNumbers>();
        Assert.Equal(1234, row.First);
        Assert.Equal(1234, row.Second);
    }

    [Fact]
    public void ReadRows_MembersOfOneTypeWithDifferentNumberStyles_UseTheirOwn()
    {
        using var importer = Helpers.GetCsvImporter("First,Second\nFF,255\n");
        importer.Configuration.RegisterClassMap<TwoNumbers>(c =>
        {
            c.Map(r => r.First).WithNumberStyle(NumberStyles.HexNumber);
            c.Map(r => r.Second);
        });

        var row = importer.ReadSheet().ReadRow<TwoNumbers>();
        Assert.Equal(255, row.First);
        Assert.Equal(255, row.Second);
    }


    public class TwoUris
    {
        public Uri First { get; set; } = default!;
        public Uri? Second { get; set; }
    }

    [Fact]
    public void ReadRows_MembersOfOneTypeWithDifferentUriKinds_UseTheirOwn()
    {
        using var importer = Helpers.GetCsvImporter("First,Second\nrelative/path,relative/path\n");
        importer.Configuration.RegisterClassMap<TwoUris>(c =>
        {
            c.Map(r => r.First).WithUriKind(UriKind.Relative);
            c.Map(r => r.Second).WithInvalidFallback(null);
        });

        var row = importer.ReadSheet().ReadRow<TwoUris>();
        Assert.Equal(new Uri("relative/path", UriKind.Relative), row.First);
        Assert.Null(row.Second);
    }

    public class TwoOffsets
    {
        public DateTimeOffset First { get; set; }
        public DateTimeOffset Second { get; set; }
    }

    [Fact]
    public void ReadRows_MembersOfOneTypeWithDifferentOffsets_UseTheirOwn()
    {
        var date = new DateTime(2024, 3, 4, 5, 6, 7);
        using var importer = XlsxBuilder.Create(["First", "Second"], [date, date]);
        importer.Configuration.RegisterClassMap<TwoOffsets>(c =>
        {
            c.Map(r => r.First).WithOffset(TimeSpan.Zero);
            c.Map(r => r.Second).WithOffset(TimeSpan.FromHours(2));
        });

        var row = importer.ReadSheet().ReadRow<TwoOffsets>();
        Assert.Equal(new DateTimeOffset(date, TimeSpan.Zero), row.First);
        Assert.Equal(new DateTimeOffset(date, TimeSpan.FromHours(2)), row.Second);
    }

#if LIBRARY_NET8
    private sealed class DoubleItConverter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context)
            => value.Text == "double" ? CellConversionResult<int>.Success(2) : CellConversionResult<int>.NotHandled;
    }

    private sealed class NullableDoubleItConverter : ICellConverter<int?>
    {
        public CellConversionResult<int?> Convert(CellValue value, CellContext context)
            => value.Text == "double" ? CellConversionResult<int?>.Success(2) : CellConversionResult<int?>.NotHandled;
    }

    public class NumberRow
    {
        public int Value { get; set; }
    }

    public class NullableNumberRow
    {
        public int? Value { get; set; }
    }

    [Fact]
    public void ReadRows_NumberStyleAfterConverter_ConvertsValuesTheConverterDoesNotHandle()
    {
        // The converter replaces the built-in conversion, and WithNumberStyle adds it back after it.
        using var importer = Helpers.GetCsvImporter("Value\ndouble\nFF\n");
        importer.Configuration.RegisterClassMap<NumberRow>(c => c.Map(r => r.Value).WithConverter(new DoubleItConverter()).WithNumberStyle(NumberStyles.HexNumber));
        Assert.Equal([2, 255], importer.ReadSheet().ReadRows<NumberRow>().Select(r => r.Value));
    }

    [Fact]
    public void ReadRows_NullableFormatProviderAfterConverter_ConvertsValuesTheConverterDoesNotHandle()
    {
        using var importer = Helpers.GetCsvImporter("Value\ndouble\n\"1.234\"\n");
        importer.Configuration.RegisterClassMap<NullableNumberRow>(c => c.Map(r => r.Value).WithConverter(new NullableDoubleItConverter()).WithFormatProvider(CultureInfo.GetCultureInfo("de-DE")));
        Assert.Equal([2, 1234], importer.ReadSheet().ReadRows<NullableNumberRow>().Select(r => r.Value));
    }

    public class ComplexRow
    {
        public Complex Value { get; set; }
    }

    [Fact]
    public void ReadRows_ComplexWithExponent_ThrowsExcelMappingException()
    {
        // Complex numbers keep Complex.Parse's default style, without exponents.
        using var importer = Helpers.GetCsvImporter("Value\n1E5\n");
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRow<ComplexRow>());
    }

    public class HalfRow
    {
        public Half Value { get; set; }
    }

    [Fact]
    public void ReadRows_HalfWithExponent_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1E2\n");
        Assert.Equal((Half)100, importer.ReadSheet().ReadRow<HalfRow>().Value);
    }

    public class BigIntegerRow
    {
        public BigInteger Value { get; set; }
    }

    [Fact]
    public void ReadRows_BigIntegerWithFraction_ThrowsExcelMappingException()
    {
        // BigInteger only has whole numbers, so a fraction is invalid rather than truncated.
        using var importer = XlsxBuilder.Create(["Value"], [2.5]);
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRow<BigIntegerRow>());
    }
#endif

    private sealed class UriConverter : ICellConverter<Uri?>
    {
        public CellConversionResult<Uri?> Convert(CellValue value, CellContext context)
            => value.Text == "home" ? CellConversionResult<Uri?>.Success(new Uri("https://example.com")) : CellConversionResult<Uri?>.NotHandled;
    }

    public class UriRow
    {
        public Uri? Value { get; set; }
    }

    [Fact]
    public void ReadRows_UriKindAfterConverter_ConvertsValuesTheConverterDoesNotHandle()
    {
        using var importer = Helpers.GetCsvImporter("Value\nhome\nrelative/path\n");
        importer.Configuration.RegisterClassMap<UriRow>(c => c.Map(r => r.Value).WithConverter(new UriConverter()).WithUriKind(UriKind.Relative));
        Assert.Equal([new Uri("https://example.com"), new Uri("relative/path", UriKind.Relative)], importer.ReadSheet().ReadRows<UriRow>().Select(r => r.Value));
    }
}
