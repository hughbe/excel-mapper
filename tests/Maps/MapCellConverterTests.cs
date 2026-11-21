using System.Globalization;
using System.Linq;
using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Tests;

/// <summary>
/// Converting cells with an <see cref="ICellConverter{T}"/>, set with WithConverter or
/// <see cref="ExcelConverterAttribute"/>.
/// </summary>
public class MapCellConverterTests
{
    [Fact]
    public void ReadRows_WithConverter_ConvertsEveryValue()
    {
        using var importer = Helpers.GetCsvImporter("Value\n3\n1.5k\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithConverter(new ThousandsConverter()));
        Assert.Equal([3, 1500], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
    }

    [Fact]
    public void ReadRows_Attribute_ConvertsEveryValue()
    {
        using var importer = Helpers.GetCsvImporter("Value,Nullable,Values\n3,2k,1;2k\n");
        var row = importer.ReadSheet().ReadRows<AttributeRow>().Single();
        Assert.Equal(3, row.Value);
        Assert.Equal(2000, row.Nullable);
        Assert.Equal([1, 2000], row.Values);
    }

    [Fact]
    public void ReadRows_AttributeWithConstructorArguments_PassesArguments()
    {
        using var importer = Helpers.GetCsvImporter("Value\n3\n");
        Assert.Equal(30, importer.ReadSheet().ReadRows<MultiplierRow>().Single().Value);
    }

    [Fact]
    public void ReadRows_AttributesChain_NotHandledTriesNext()
    {
        using var importer = Helpers.GetCsvImporter("Value\n3\nn/a\n");
        Assert.Equal([3, -1], importer.ReadSheet().ReadRows<ChainRow>().Select(r => r.Value));
    }

    [Fact]
    public void ReadRows_NoConverterHandles_IsInvalid()
    {
        using var importer = Helpers.GetCsvImporter("Value\nx\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<ChainRow>().ToArray());
        Assert.Equal("A2", exception.CellAddress);
        Assert.Null(exception.InnerException);
    }

    [Fact]
    public void ReadRows_AttributeOnCollectionWithWrongElementType_Throws()
    {
        using var importer = Helpers.GetCsvImporter("Values\n1;2\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<WrongElementTypeRow>().ToArray());
        Assert.Contains("ICellConverter", exception.Message);
    }

    public class WrongElementTypeRow
    {
        [ExcelSeparators(';')]
        [ExcelConverter(typeof(StringConverter))]
        public int[] Values { get; set; } = [];
    }

    [Fact]
    public void ReadRows_AttributeWrongType_Throws()
    {
        using var importer = Helpers.GetCsvImporter("Value\n3\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<WrongTypeRow>().ToArray());
        Assert.Contains("ICellConverter", exception.Message);
    }

    [Fact]
    public void ReadRows_NullableWithConverterOfUnderlyingType_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Value,Other\n2k,1\n,2\n");
        importer.Configuration.RegisterClassMap<NullableRow>(c => c.Map(r => r.Value).WithConverter(new ThousandsConverter()));
        Assert.Equal([2000, null], importer.ReadSheet().ReadRows<NullableRow>().Select(r => r.Value));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ReadRows_InvalidWithException_ThrowsWithInnerException(bool lazy)
    {
        using var importer = Helpers.GetCsvImporter("Value\n3\nx\n");
        var converter = new ExceptionConverter(lazy);
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithConverter(converter));
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<IntRow>().ToArray());
        Assert.Equal("A3", exception.CellAddress);
        Assert.IsType<FormatException>(exception.InnerException);
        Assert.Equal(1, converter.ExceptionsCreated);
    }

    [Fact]
    public void ReadRows_InvalidWithFixedFallback_DoesNotCreateException()
    {
        using var importer = Helpers.GetCsvImporter("Value\nx\ny\n");
        var converter = new ExceptionConverter(lazy: true);
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithConverter(converter).WithInvalidFallback(-1));
        Assert.Equal([-1, -1], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
        Assert.Equal(0, converter.ExceptionsCreated);
    }

    [Fact]
    public void ReadRowResults_Invalid_ReportsError()
    {
        using var importer = Helpers.GetCsvImporter("Value\nx\n4\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithConverter(new ExceptionConverter(lazy: false)));
        var results = importer.ReadSheet().ReadRowResults<IntRow>().ToArray();
        Assert.False(results[0].IsValid);
        Assert.IsType<FormatException>(Assert.Single(results[0].Errors).InnerException);
        Assert.Equal(4, results[1].Value!.Value);
    }

    [Fact]
    public void ReadRows_EmptyCell_UsesEmptyFallbackWithoutConverter()
    {
        using var importer = Helpers.GetCsvImporter("Value,Other\n,1\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithConverter(new ThrowingConverter()).WithEmptyFallback(7));
        Assert.Equal(7, importer.ReadSheet().ReadRows<IntRow>().Single().Value);
    }

    [Fact]
    public void ReadRows_Context_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Other,Value\nx,1\ny,2\n");
        var culture = new CultureInfo("fr-FR");
        importer.Configuration.Culture = culture;
        var converter = new ContextConverter();
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithConverter(converter));
        importer.ReadSheet().ReadRows<IntRow>().ToArray();

        Assert.Equal(["B2", "B3"], converter.Contexts.Select(c => c.CellAddress));
        Assert.Equal([1, 2], converter.Contexts.Select(c => c.RowIndex));
        Assert.All(converter.Contexts, c => Assert.Equal(1, c.ColumnIndex));
        Assert.All(converter.Contexts, c => Assert.Equal(nameof(IntRow.Value), c.Member!.Name));
        Assert.All(converter.Contexts, c => Assert.Same(culture, c.FormatProvider));
    }

    [Fact]
    public void ReadRows_WithOtherMapperAfterConverter_PassesContext()
    {
        // A mapper that isn't a converter means values go through the pipeline as objects: the
        // converter still gets the cell's context.
        using var importer = Helpers.GetCsvImporter("Value\n1\nskip\n");
        var converter = new ContextConverter();
        importer.Configuration.RegisterClassMap<IntRow>(c =>
        {
            var map = c.Map(r => r.Value).WithConverter(converter);
            map.Mappers.Add(new SkipMapper());
        });
        Assert.Equal([1, 0], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
        Assert.Equal(["A2", "A3"], converter.Contexts.Select(c => c.CellAddress));
        Assert.All(converter.Contexts, c => Assert.Equal(nameof(IntRow.Value), c.Member!.Name));
    }

    [Fact]
    public void ReadRows_ElementsAndDictionaryValues_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Values,A,B\n1;2k,3k,4\n");
        importer.Configuration.RegisterClassMap<CollectionRow>(c =>
        {
            c.Map(r => r.Values).WithSeparators(';').WithElementMap(e => e.WithConverter(new ThousandsConverter()));
            c.Map(r => r.Columns).WithColumnNames("A", "B").WithValueMap(v => v.WithConverter(new ThousandsConverter()));
        });
        var row = importer.ReadSheet().ReadRows<CollectionRow>().Single();
        Assert.Equal([1, 2000], row.Values);
        Assert.Equal(new Dictionary<string, int> { ["A"] = 3000, ["B"] = 4 }, row.Columns);
    }

    [Fact]
    public void ReadRows_RecordWithAttribute_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,Value\na,2k\n");
        Assert.Equal(new RecordRow("a", 2000), importer.ReadSheet().ReadRows<RecordRow>().Single());
    }

    [Fact]
    public void ReadRows_TypedCellXlsx_ReceivesNumber()
    {
        using var importer = XlsxBuilder.Create(["Value"], [2.5]);
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithConverter(new RoundingConverter()));
        Assert.Equal(3, importer.ReadSheet().ReadRows<IntRow>().Single().Value);
    }

    public class IntRow
    {
        public int Value { get; set; }
    }

    public class NullableRow
    {
        public int? Value { get; set; }
        public int Other { get; set; }
    }

    public class AttributeRow
    {
        [ExcelConverter(typeof(ThousandsConverter))]
        public int Value { get; set; }

        [ExcelConverter(typeof(ThousandsConverter))]
        public int? Nullable { get; set; }

        [ExcelSeparators(';')]
        [ExcelConverter(typeof(ThousandsConverter))]
        public int[] Values { get; set; } = [];
    }

    public class MultiplierRow
    {
        [ExcelConverter(typeof(MultiplyingConverter), ConstructorArguments = [10])]
        public int Value { get; set; }
    }

    public class ChainRow
    {
        [ExcelConverter(typeof(NumberOnlyConverter))]
        [ExcelConverter(typeof(NotApplicableConverter))]
        public int Value { get; set; }
    }

    public class WrongTypeRow
    {
        [ExcelConverter(typeof(StringConverter))]
        public int Value { get; set; }
    }

    public class CollectionRow
    {
        public int[] Values { get; set; } = [];
        public Dictionary<string, int> Columns { get; set; } = [];
    }

    public record RecordRow(string Name, [property: ExcelConverter(typeof(ThousandsConverter))] int Value);

    public class ThousandsConverter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context)
        {
            var text = value.Text!;
            return text.EndsWith("k", StringComparison.Ordinal)
                ? CellConversionResult<int>.Success((int)(double.Parse(text[..^1], CultureInfo.InvariantCulture) * 1000))
                : CellConversionResult<int>.Success(int.Parse(text, CultureInfo.InvariantCulture));
        }
    }

    public class MultiplyingConverter(int factor) : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context)
            => CellConversionResult<int>.Success(int.Parse(value.Text!, CultureInfo.InvariantCulture) * factor);
    }

    public class NumberOnlyConverter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context)
            => int.TryParse(value.Text, out var result) ? CellConversionResult<int>.Success(result) : CellConversionResult<int>.NotHandled;
    }

    public class NotApplicableConverter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context)
            => value.Text == "n/a" ? CellConversionResult<int>.Success(-1) : CellConversionResult<int>.NotHandled;
    }

    public class StringConverter : ICellConverter<string>
    {
        public CellConversionResult<string> Convert(CellValue value, CellContext context) => CellConversionResult<string>.Success(value.Text!);
    }

    private sealed class ExceptionConverter(bool lazy) : ICellConverter<int>
    {
        public int ExceptionsCreated { get; private set; }

        public CellConversionResult<int> Convert(CellValue value, CellContext context)
        {
            if (int.TryParse(value.Text, out var result))
            {
                return CellConversionResult<int>.Success(result);
            }

            if (lazy)
            {
                return CellConversionResult<int>.Invalid(() =>
                {
                    ExceptionsCreated++;
                    return new FormatException($"\"{value.Text}\" isn't a number.");
                });
            }

            ExceptionsCreated++;
            return CellConversionResult<int>.Invalid(new FormatException($"\"{value.Text}\" isn't a number."));
        }
    }

    private sealed class ThrowingConverter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context) => throw new InvalidOperationException("Not called for empty cells.");
    }

    private sealed class ContextConverter : ICellConverter<int>
    {
        public List<CellContext> Contexts { get; } = [];

        public CellConversionResult<int> Convert(CellValue value, CellContext context)
        {
            Contexts.Add(context);
            return int.TryParse(value.Text, out var result) ? CellConversionResult<int>.Success(result) : CellConversionResult<int>.NotHandled;
        }
    }

    private sealed class SkipMapper : ICellMapper
    {
        public CellMapperResult Map(ReadCellResult readResult) => readResult.GetString() == "skip" ? CellMapperResult.Success(0) : CellMapperResult.Ignore();
    }

    private sealed class RoundingConverter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context)
            => value.Kind == CellKind.Number ? CellConversionResult<int>.Success((int)Math.Round(value.Number, MidpointRounding.AwayFromZero)) : CellConversionResult<int>.NotHandled;
    }
}
