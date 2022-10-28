using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;
 
public class OneToOneMapTTests
{
    [Fact]
    public void Ctor_ICellReader()
    {
        var reader = new ColumnNameValueReader("Column");
        var map = new SubOneToOneMap<int>(reader);
        Assert.Same(reader, map.CellReader);
        Assert.Empty(map.Mappers);
        Assert.False(map.Optional);
    }

    [Fact]
    public void Ctor_NullReader_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("reader", () => new SubOneToOneMap<int>(null));
    }

    public static IEnumerable<object[]> CellReader_Set_TestData()
    {
        yield return new object[] { new ColumnNameValueReader("Column") };
    }

    [Theory]
    [MemberData(nameof(CellReader_Set_TestData))]
    public void CellReader_SetValid_GetReturnsExpected(ICellReader value)
    {
        var reader = new ColumnNameValueReader("Column");
        var map = new SubOneToOneMap<int>(reader)
        {
            CellReader = value
        };
        Assert.Same(value, map.CellReader);

        // Set same.
        map.CellReader = value;
        Assert.Same(value, map.CellReader);
    }

    [Fact]
    public void CellReader_SetNull_ThrowsArgumentNullException()
    {
        var reader = new ColumnNameValueReader("Column");
        var map = new SubOneToOneMap<int>(reader);

        Assert.Throws<ArgumentNullException>("value", () => map.CellReader = null);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Optional_Set_GetReturnsExpected(bool value)
    {
        var reader = new ColumnNameValueReader("Column");
        var map = new SubOneToOneMap<int>(reader)
        {
            Optional = value
        };
        Assert.Equal(value, map.Optional);

        // Set same.
        map.Optional = value;
        Assert.Equal(value, map.Optional);

        // Set different.
        map.Optional = !value;
        Assert.Equal(!value, map.Optional);
    }

    private class TestClass
    {
        public string Value { get; set; }
    }

    private class SubOneToOneMap<T> : OneToOneMap<T>
    {
        public SubOneToOneMap(ICellReader reader) : base(reader)
        {
        }
    }
}
