namespace ExcelMapper.Tests;

public class ExcelColumnIndicesAttributeTests
{
    public static IEnumerable<object[]> Ctor_ParamsInt_TestData()
    {
        yield return new object[] { new int[] { 0 } };
        yield return new object[] { new int[] { 0, 0 } };
        yield return new object[] { new int[] { 0, 1 } };
    }

    [Theory]
    [MemberData(nameof(Ctor_ParamsInt_TestData))]
    public void Ctor_ParamsInt(int[] columnIndices)
    {
        var reader = new ExcelColumnIndicesAttribute(columnIndices);
        Assert.Same(columnIndices, reader.Indices);
    }

    [Fact]
    public void Ctor_NullColumnIndices_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnIndices", () => new ExcelColumnIndicesAttribute(null!));
    }

    [Fact]
    public void Ctor_EmptyColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnIndices", () => new ExcelColumnIndicesAttribute([]));
    }

    [Fact]
    public void Ctor_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => new ExcelColumnIndicesAttribute([-1]));
    }
    public static IEnumerable<object[]> Indices_Set_TestData()
    {
        yield return new object[] { new int[] { 0 } };
        yield return new object[] { new int[] { 0, 0 } };
        yield return new object[] { new int[] { 0, 1 } };
    }

    [Theory]
    [MemberData(nameof(Indices_Set_TestData))]
    public void Indices_Set_GetReturnsExpected(int[] value)
    {
        var reader = new ExcelColumnIndicesAttribute(1)
        {
            Indices = value
        };
        Assert.Same(value, reader.Indices);

        // Set same.
        reader.Indices = value;
        Assert.Same(value, reader.Indices);
    }

    [Fact]
    public void Indices_SetNullValue_ThrowsArgumentNullException()
    {
        var attribute = new ExcelColumnIndicesAttribute(1);
        Assert.Throws<ArgumentNullException>("value", () => attribute.Indices = null!);
    }

    [Fact]
    public void Indices_SetEmptyValue_ThrowsArgumentException()
    {
        var attribute = new ExcelColumnIndicesAttribute(1);
        Assert.Throws<ArgumentException>("value", () => attribute.Indices = []);
    }

    [Fact]
    public void Indices_SetNegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
    {
        var attribute = new ExcelColumnIndicesAttribute(1);
        Assert.Throws<ArgumentOutOfRangeException>("value", () => attribute.Indices = [-1]);
    }
}
