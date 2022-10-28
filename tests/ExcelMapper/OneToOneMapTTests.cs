using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;
 
public class OneToOneMapTTests
{
    [Fact]
    public void Ctor_IReader()
    {
        var reader = new ColumnNameValueReader("Column");
        var map = new SubOneToOneMap<int>(reader);
        Assert.Same(reader, map.Reader);
        Assert.Empty(map.Mappers);
        Assert.False(map.Optional);
    }

    [Fact]
    public void Ctor_NullReader_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("reader", () => new SubOneToOneMap<int>(null));
    }

    public static IEnumerable<object[]> Reader_Set_TestData()
    {
        yield return new object[] { new ColumnNameValueReader("Column") };
    }

    [Theory]
    [MemberData(nameof(Reader_Set_TestData))]
    public void Reader_SetValid_GetReturnsExpected(ICellReader value)
    {
        var reader = new ColumnNameValueReader("Column");
        var map = new SubOneToOneMap<int>(reader)
        {
            Reader = value
        };
        Assert.Same(value, map.Reader);

        // Set same.
        map.Reader = value;
        Assert.Same(value, map.Reader);
    }

    [Fact]
    public void Reader_SetNull_ThrowsArgumentNullException()
    {
        var reader = new ColumnNameValueReader("Column");
        var map = new SubOneToOneMap<int>(reader);

        Assert.Throws<ArgumentNullException>("value", () => map.Reader = null);
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
