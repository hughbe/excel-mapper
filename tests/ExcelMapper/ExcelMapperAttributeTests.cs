using ExcelMapper.Abstractions;
using ExcelMapper.Mappers;

namespace ExcelMapper.Tests;

public class ExcelMapperAttributeTests
{
    [Theory]
    [InlineData(typeof(CellMapper))]
    [InlineData(typeof(BoolMapper))]
    [InlineData(typeof(NoConstructorCellMapper))]
    public void Ctor_Type(Type mapperType)
    {
        var attribute = new ExcelMapperAttribute(mapperType);
        Assert.Same(mapperType, attribute.Type);
        Assert.Null(attribute.ConstructorArguments);
    }

    [Fact]
    public void Ctor_NullMapperType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("mapperType", () => new ExcelMapperAttribute(null!));
    }


    [Theory]
    [InlineData(typeof(ICellMapper))]
    [InlineData(typeof(ISubCellMapper))]
    [InlineData(typeof(AbstractCellMapper))]
    [InlineData(typeof(int))]
    [InlineData(typeof(object))]
    [InlineData(typeof(ExcelMapperAttributeTests))]
    public void Ctor_InvalidMapperType_ThrowsArgumentException(Type mapperType)
    {
        Assert.Throws<ArgumentException>("mapperType", () => new ExcelMapperAttribute(mapperType));
    }

    private interface ISubCellMapper : ICellMapper
    {
    }

    private abstract class AbstractCellMapper : ICellMapper
    {
        public abstract CellMapperResult Map(ReadCellResult readResult);
    }

    private class NoConstructorCellMapper : ICellMapper
    {
        private NoConstructorCellMapper()
        {
        }

        public CellMapperResult Map(ReadCellResult readResult) => throw new NotImplementedException();
    }

    public static IEnumerable<object?[]> ConstructorArguments_Set_TestData()
    {
        yield return new object?[] { null };
        yield return new object[] { new object[0] };
        yield return new object[] { new object?[] { "Value", null } };
    }

    [Theory]
    [MemberData(nameof(ConstructorArguments_Set_TestData))]
    public void ConstructorArguments_Set_GetReturnsExpected(object?[]? value)
    {
        var attribute = new ExcelMapperAttribute(typeof(CellMapper))
        {
            ConstructorArguments = value
        };
        Assert.Same(value, attribute.ConstructorArguments);
        
        // Set.
        attribute.ConstructorArguments = value;
        Assert.Same(value, attribute.ConstructorArguments);
    }

    private class CellMapper : ICellMapper
    {
        public CellMapperResult Map(ReadCellResult readResult) => throw new NotImplementedException();
    }
}
