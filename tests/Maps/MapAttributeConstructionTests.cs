namespace ExcelMapper.Tests;

/// <summary>
/// Attributes such as <see cref="ExcelConverterAttribute"/> give a type and its constructor
/// arguments. If the type can't be created with them, auto-mapping throws an
/// <see cref="ExcelMappingException"/> naming the type, the attribute and the member.
/// </summary>
public class MapAttributeConstructionTests
{
    public class TakesNumber : ICellConverter<int>, ICellFallback<int>, ICellTransformer, IColumnSelector
    {
        public TakesNumber(int number) => Number = number;

        public int Number { get; }

        public CellConversionResult<int> Convert(CellValue value, CellContext context) => CellConversionResult<int>.Success(Number);

        public int Fallback(CellValue cell, CellContext context, Exception? exception) => Number;

        public CellValue Transform(CellValue value, CellContext context) => value;

        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context) => [0];

        public string Describe(ColumnSelectorContext context) => "the first column";
    }

    public class Throws : ICellConverter<int>
    {
        public Throws() => throw new InvalidOperationException("Not today.");

        public CellConversionResult<int> Convert(CellValue value, CellContext context) => CellConversionResult<int>.NotHandled;
    }

    public class ConverterWithWrongArguments
    {
        [ExcelConverter(typeof(TakesNumber), ConstructorArguments = ["one"])]
        public int Value { get; set; }
    }

    public class ConverterWithoutArguments
    {
        [ExcelConverter(typeof(TakesNumber))]
        public int Value { get; set; }
    }

    public class ConverterThatThrows
    {
        [ExcelConverter(typeof(Throws))]
        public int Value { get; set; }
    }

    public class EmptyFallbackWithWrongArguments
    {
        [ExcelEmptyFallback(typeof(TakesNumber), ConstructorArguments = ["one"])]
        public int Value { get; set; }
    }

    public class InvalidFallbackWithWrongArguments
    {
        [ExcelInvalidFallback(typeof(TakesNumber), ConstructorArguments = ["one"])]
        public int Value { get; set; }
    }

    public class TransformerWithWrongArguments
    {
        [ExcelTransformer(typeof(TakesNumber), ConstructorArguments = ["one"])]
        public string Value { get; set; } = default!;
    }

    public class SelectorWithWrongArguments
    {
        [ExcelColumnMatching(typeof(TakesNumber), ConstructorArguments = ["one"])]
        public int Value { get; set; }
    }

    public class SelectorsWithWrongArguments
    {
        [ExcelColumnsMatching(typeof(TakesNumber), ConstructorArguments = ["one"])]
        public int[] Value { get; set; } = default!;
    }

    private static ExcelMappingException ReadRow<T>()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        return Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRow<T>());
    }

    [Theory]
    [InlineData(typeof(ConverterWithWrongArguments), "ExcelConverterAttribute")]
    [InlineData(typeof(ConverterWithoutArguments), "ExcelConverterAttribute")]
    [InlineData(typeof(EmptyFallbackWithWrongArguments), "ExcelEmptyFallbackAttribute")]
    [InlineData(typeof(InvalidFallbackWithWrongArguments), "ExcelInvalidFallbackAttribute")]
    [InlineData(typeof(TransformerWithWrongArguments), "ExcelTransformerAttribute")]
    [InlineData(typeof(SelectorWithWrongArguments), "ExcelColumnMatchingAttribute")]
    [InlineData(typeof(SelectorsWithWrongArguments), "ExcelColumnsMatchingAttribute")]
    public void ReadRow_AttributeTypeCantBeCreated_ThrowsExcelMappingException(Type rowType, string attributeName)
    {
        var method = typeof(MapAttributeConstructionTests).GetMethod(nameof(ReadRow), System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!.MakeGenericMethod(rowType);
        var exception = (ExcelMappingException)method.Invoke(null, null)!;
        Assert.Contains(typeof(TakesNumber).ToString(), exception.Message);
        Assert.Contains(attributeName, exception.Message);
        Assert.Contains("member \"Value\"", exception.Message);
        Assert.IsType<MissingMethodException>(exception.InnerException);
    }

    [Fact]
    public void ReadRow_AttributeTypeConstructorThrows_ThrowsExcelMappingExceptionWithItsException()
    {
        var exception = ReadRow<ConverterThatThrows>();
        Assert.Contains("given by ExcelConverterAttribute on member \"Value\": Not today.", exception.Message);
        Assert.IsType<InvalidOperationException>(exception.InnerException);
    }
}
