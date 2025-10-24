using ExcelMapper.Abstractions;
using ExcelMapper.Transformers;

namespace ExcelMapper.Tests;

public class ExcelTransformerAttributeTests
{
    [Theory]
    [InlineData(typeof(CellTransformer))]
    [InlineData(typeof(TrimStringCellTransformer))]
    [InlineData(typeof(NoConstructorCellTransformer))]
    public void Ctor_Type(Type transformerType)
    {
        var attribute = new ExcelTransformerAttribute(transformerType);
        Assert.Same(transformerType, attribute.Type);
        Assert.Null(attribute.ConstructorArguments);
    }

    [Fact]
    public void Ctor_NullTransformerType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("transformerType", () => new ExcelTransformerAttribute(null!));
    }


    [Theory]
    [InlineData(typeof(ICellTransformer))]
    [InlineData(typeof(ISubCellTransformer))]
    [InlineData(typeof(AbstractCellTransformer))]
    [InlineData(typeof(int))]
    [InlineData(typeof(object))]
    [InlineData(typeof(ExcelTransformerAttributeTests))]
    public void Ctor_InvalidTransformerType_ThrowsArgumentException(Type transformerType)
    {
        Assert.Throws<ArgumentException>("transformerType", () => new ExcelTransformerAttribute(transformerType));
    }

    private interface ISubCellTransformer : ICellTransformer
    {
    }

    private abstract class AbstractCellTransformer : ICellTransformer
    {
        public abstract string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult);
    }

    private class NoConstructorCellTransformer : ICellTransformer
    {
        private NoConstructorCellTransformer()
        {
        }

        public string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
            => throw new NotImplementedException();
    }

    public static IEnumerable<object?[]> ConstructorArguments_Set_TestData()
    {
        yield return new object?[] { null };
        yield return new object[] { Array.Empty<object>() };
        yield return new object[] { new object?[] { "Value", null } };
    }

    [Theory]
    [MemberData(nameof(ConstructorArguments_Set_TestData))]
    public void ConstructorArguments_Set_GetReturnsExpected(object?[]? value)
    {
        var attribute = new ExcelTransformerAttribute(typeof(CellTransformer))
        {
            ConstructorArguments = value
        };
        Assert.Same(value, attribute.ConstructorArguments);
        
        // Set.
        attribute.ConstructorArguments = value;
        Assert.Same(value, attribute.ConstructorArguments);
    }

    private class CellTransformer : ICellTransformer
    {
        public string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
            => throw new NotImplementedException();
    }
}
