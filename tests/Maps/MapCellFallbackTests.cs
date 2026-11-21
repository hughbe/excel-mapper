using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// Empty and invalid cells with an <see cref="ICellFallback{T}"/>, set with WithEmptyFallback,
/// WithInvalidFallback, <see cref="ExcelEmptyFallbackAttribute"/> or <see cref="ExcelInvalidFallbackAttribute"/>.
/// </summary>
public class MapCellFallbackTests
{
    [Fact]
    public void ReadRows_WithEmptyFallback_UsedForEmptyCells()
    {
        var fallback = new RecordingFallback(-1);
        using var importer = Helpers.GetCsvImporter("Value,Other\n,1\n5,2\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithEmptyFallback(fallback));
        Assert.Equal([-1, 5], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
        var call = Assert.Single(fallback.Calls);
        Assert.Equal("A2", call.Address);
        Assert.True(call.IsEmpty);
        Assert.Null(call.Exception);
        Assert.Equal(nameof(IntRow.Value), call.Member);
    }

    [Fact]
    public void ReadRows_WithInvalidFallback_UsedForInvalidCellsWithException()
    {
        var fallback = new RecordingFallback(-2);
        using var importer = Helpers.GetCsvImporter("Value,Other\nx,1\n5,2\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithInvalidFallback(fallback));
        Assert.Equal([-2, 5], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
        var call = Assert.Single(fallback.Calls);
        Assert.Equal("A2", call.Address);
        Assert.Equal("x", call.Text);
        Assert.IsType<FormatException>(call.Exception);
    }

    [Fact]
    public void ReadRows_Attributes_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Value,Nullable,Other\n,x,1\nx,,2\n");
        var rows = importer.ReadSheet().ReadRows<AttributeRow>().ToArray();
        Assert.Equal(100, rows[0].Value);
        Assert.Equal(200, rows[0].Nullable);
        Assert.Equal(200, rows[1].Value);
        Assert.Equal(100, rows[1].Nullable);
    }

    [Fact]
    public void ReadRows_NullableWithNullableFallback_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Value,Other\n,1\nx,2\n");
        importer.Configuration.RegisterClassMap<NullableRow>(c => c.Map(r => r.Value).WithEmptyFallback(new NullableFallback(7)).WithInvalidFallback(new NullableFallback(8)));
        Assert.Equal([7, 8], importer.ReadSheet().ReadRows<NullableRow>().Select(r => r.Value));
    }

    [Fact]
    public void ReadRows_FallbackThrows_ReportsError()
    {
        using var importer = Helpers.GetCsvImporter("Value,Other\nx,1\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithInvalidFallback(new ThrowingFallback()));
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<IntRow>().ToArray());
        Assert.Equal("Rejected.", exception.Message);
    }

    [Fact]
    public void ReadRowResults_FallbackThrows_RecordsError()
    {
        using var importer = Helpers.GetCsvImporter("Value,Other\nx,1\n4,2\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithInvalidFallback(new ThrowingFallback()));
        var results = importer.ReadSheet().ReadRowResults<IntRow>().ToArray();
        Assert.False(results[0].IsValid);
        Assert.Equal("Rejected.", Assert.Single(results[0].Errors).Message);
        Assert.Equal(4, results[1].Value!.Value);
    }

    [Fact]
    public void ReadRows_WithConverterThatReturnsException_FallbackReceivesIt()
    {
        var fallback = new RecordingFallback(0);
        using var importer = Helpers.GetCsvImporter("Value,Other\nx,1\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value)
            .WithConverter(new InvalidConverter())
            .WithInvalidFallback(fallback));
        importer.ReadSheet().ReadRows<IntRow>().ToArray();
        Assert.Equal("Not a number.", Assert.Single(fallback.Calls).Exception!.Message);
    }

    [Fact]
    public void ReadRows_ObjectPipeline_UsesFallback()
    {
        using var importer = Helpers.GetCsvImporter("Value,Other\nx,1\n");
        importer.Configuration.RegisterClassMap<IntRow>(c =>
        {
            // A mapping dictionary isn't a converter, so values go through the pipeline as objects.
            c.Map(r => r.Value).WithInvalidFallback(new FixedIntFallback(9)).WithMapping(new Dictionary<string, int> { ["y"] = 1 });
        });
        Assert.Equal(9, importer.ReadSheet().ReadRows<IntRow>().Single().Value);
    }

    [Fact]
    public void ReadRows_ElementsAndDictionaryValues_UseFallbacks()
    {
        using var importer = Helpers.GetCsvImporter("Values,A,B\n1;x;,2,y\n");
        importer.Configuration.RegisterClassMap<CollectionRow>(c =>
        {
            c.Map(r => r.Values).WithSeparators(';').WithElementMap(e => e.WithInvalidFallback(new FixedIntFallback(-1)).WithEmptyFallback(new FixedIntFallback(0)));
            c.Map(r => r.Columns).WithColumnNames("A", "B").WithValueMap(v => v.WithInvalidFallback(new FixedIntFallback(-1)));
        });
        var row = importer.ReadSheet().ReadRows<CollectionRow>().Single();
        Assert.Equal([1, -1, 0], row.Values);
        Assert.Equal(new Dictionary<string, int> { ["A"] = 2, ["B"] = -1 }, row.Columns);
    }

    [Fact]
    public void ReadRows_AttributeWrongType_Throws()
    {
        using var importer = Helpers.GetCsvImporter("Value,Other\n,1\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<WrongTypeRow>().ToArray());
        Assert.Contains("ICellFallback", exception.Message);
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
        [ExcelEmptyFallback(typeof(FixedFallback), ConstructorArguments = [100])]
        [ExcelInvalidFallback(typeof(FixedFallback), ConstructorArguments = [200])]
        public int Value { get; set; }

        [ExcelEmptyFallback(typeof(FixedFallback), ConstructorArguments = [100])]
        [ExcelInvalidFallback(typeof(FixedFallback), ConstructorArguments = [200])]
        public int? Nullable { get; set; }
    }

    public class WrongTypeRow
    {
        [ExcelEmptyFallback(typeof(StringFallback))]
        public int Value { get; set; }
        public int Other { get; set; }
    }

    public class CollectionRow
    {
        public int[] Values { get; set; } = [];
        public Dictionary<string, int> Columns { get; set; } = [];
    }

    public class FixedFallback(int value) : ICellFallback<int>
    {
        public int Fallback(CellValue cell, CellContext context, Exception? exception) => value;
    }

    public class NullableFallback(int value) : ICellFallback<int?>
    {
        public int? Fallback(CellValue cell, CellContext context, Exception? exception) => value;
    }

    public class FixedIntFallback(int value) : ICellFallback<int>
    {
        public int Fallback(CellValue cell, CellContext context, Exception? exception) => value;
    }

    public class StringFallback : ICellFallback<string>
    {
        public string Fallback(CellValue cell, CellContext context, Exception? exception) => "";
    }

    private sealed class ThrowingFallback : ICellFallback<int>
    {
        public int Fallback(CellValue cell, CellContext context, Exception? exception) => throw new ExcelMappingException("Rejected.");
    }

    private sealed class InvalidConverter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context) => CellConversionResult<int>.Invalid(() => new FormatException("Not a number."));
    }


    public record Call(string Address, string? Text, bool IsEmpty, Exception? Exception, string? Member);

    private sealed class RecordingFallback(int value) : ICellFallback<int>
    {
        public List<Call> Calls { get; } = [];

        public int Fallback(CellValue cell, CellContext context, Exception? exception)
        {
            Calls.Add(new Call(context.CellAddress, cell.Text, cell.IsEmpty, exception, context.Member?.Name));
            return value;
        }
    }
}
