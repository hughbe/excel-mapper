namespace ExcelMapper.Readers.Tests;

public class StringSplitReaderFactoryTests
{
    [Fact]
    public void Ctor_ICellReaderFactory()
    {
        var valueFactory = new ColumnNameReaderFactory("ColumnName");
        var factory = new StringSplitReaderFactory(valueFactory);
        Assert.Same(valueFactory, factory.ReaderFactory);
        Assert.Equal(StringSplitOptions.None, factory.Options);
        Assert.Equal([","], factory.Separators);
        Assert.Same(factory.Separators, factory.Separators);
    }

    public static IEnumerable<object[]> Separators_Set_TestData()
    {
        yield return new object[] { new string[] { "," } };
        yield return new object[] { new string[] { ",", ";" } };
    }

    [Theory]
    [MemberData(nameof(Separators_Set_TestData))]
    public void Separators_SetValid_GetReturnsExpected(string[] value)
    {
        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"))
        {
            Separators = value
        };
        Assert.Same(value, factory.Separators);

        // Set same.
        factory.Separators = value;
        Assert.Same(value, factory.Separators);
    }

    [Fact]
    public void Separators_SetNull_ThrowsArgumentNullException()
    {
        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.Throws<ArgumentNullException>("value", () => factory.Separators = null!);
    }

    [Fact]
    public void Separators_SetEmpty_ThrowsArgumentException()
    {
        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.Throws<ArgumentException>("value", () => factory.Separators = []);
    }

    [Fact]
    public void Separators_SetNullValueInArray_ThrowsArgumentException()
    {
        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.Throws<ArgumentException>("value", () => factory.Separators = [",", null!]);
    }

    [Fact]
    public void Separators_SetEmptyValueInArray_ThrowsArgumentException()
    {
        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.Throws<ArgumentException>("value", () => factory.Separators = [",", ""]);
    }
}
