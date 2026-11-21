namespace ExcelMapper.Tests;

public class ExcelConverterAttributeTests
{
    [Fact]
    public void Ctor_Type()
    {
        var attribute = new ExcelConverterAttribute(typeof(Converter));
        Assert.Same(typeof(Converter), attribute.Type);
        Assert.Null(attribute.ConstructorArguments);
    }

    [Fact]
    public void ConstructorArguments_Set_GetReturnsExpected()
    {
        var attribute = new ExcelConverterAttribute(typeof(Converter)) { ConstructorArguments = [1] };
        Assert.Equal([1], attribute.ConstructorArguments);
    }

    [Fact]
    public void Ctor_NullConverterType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("converterType", () => new ExcelConverterAttribute(null!));
    }

    [Theory]
    [InlineData(typeof(ICellConverter<int>))]
    [InlineData(typeof(AbstractConverter))]
    [InlineData(typeof(int))]
    [InlineData(typeof(object))]
    public void Ctor_InvalidConverterType_ThrowsArgumentException(Type converterType)
    {
        Assert.Throws<ArgumentException>("converterType", () => new ExcelConverterAttribute(converterType));
    }

    private sealed class Converter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context) => CellConversionResult<int>.Success(1);
    }

    private abstract class AbstractConverter : ICellConverter<int>
    {
        public abstract CellConversionResult<int> Convert(CellValue value, CellContext context);
    }
}
