using ExcelDataReader;
using System.Diagnostics.CodeAnalysis;
using System.Data;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers.Tests;

public class ChangeTypeMapperTests
{
    [Theory]
    [InlineData(typeof(string))]
    [InlineData(typeof(int))]
    public void Ctor_Type(Type type)
    {
        var item = new ChangeTypeMapper(type);
        Assert.Equal(type, item.Type);
    }

    [Fact]
    public void Ctor_NullType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("type", () => new ChangeTypeMapper(null!));
    }

    [Fact]
    public void Ctor_TypeNotIConvertible_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("type", () => new ChangeTypeMapper(typeof(List<int>)));
    }

    [Theory]
    [InlineData(typeof(int), "1", 1)]
    public void Map_ValidStringValue_ReturnsSuccess(Type type, string stringValue, object expected)
    {
        var item = new ChangeTypeMapper(type);

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, result.Value);
        Assert.Null(result.Exception);
    }

    [Theory]
    [InlineData(typeof(uint), "-1")]
    [InlineData(typeof(uint), "abc")]
    [InlineData(typeof(uint), "")]
    [InlineData(typeof(uint), null)]
    public void Map_InvalidStringValue_ReturnsInvalid(Type type, string? stringValue)
    {
        var item = new ChangeTypeMapper(type);

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.NotNull(result.Exception);
    }

    [Theory]
    [InlineData(typeof(int), 2.7)]
    [InlineData(typeof(long), -0.5)]
    [InlineData(typeof(byte), 1.25)]
    public void Map_FractionalDoubleValueToInteger_ReturnsInvalid(Type type, double value)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => value
        };
        var item = new ChangeTypeMapper(type);

        var result = item.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.IsType<FormatException>(result.Exception);
    }

    [Theory]
    [InlineData(typeof(int), 2.0, 2)]
    [InlineData(typeof(double), 2.7, 2.7)]
    public void Map_DoubleValue_ReturnsSuccess(Type type, double value, object expected)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => value
        };
        var item = new ChangeTypeMapper(type);

        var result = item.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, result.Value);
    }

    private class MockExcelDataReader : IExcelDataReader
    {
        public object this[int i] => throw new NotImplementedException();

        public object this[string name] => throw new NotImplementedException();

        public string Name => throw new NotImplementedException();

        public string CodeName => throw new NotImplementedException();

        public string VisibleState => throw new NotImplementedException();

        public int ActiveSheet => throw new NotImplementedException();

        public bool IsActiveSheet => throw new NotImplementedException();

        public HeaderFooter HeaderFooter => throw new NotImplementedException();

        public CellRange[] MergeCells => throw new NotImplementedException();

        public int ResultsCount => throw new NotImplementedException();

        public int RowCount => throw new NotImplementedException();

        public double RowHeight => throw new NotImplementedException();

        public CellRange Dimension => throw new NotImplementedException();

        public int Depth => throw new NotImplementedException();

        public bool IsClosed => throw new NotImplementedException();

        public int RecordsAffected => throw new NotImplementedException();

        public int FieldCount => throw new NotImplementedException();

        public void Close() => throw new NotImplementedException();

        public void Dispose() => throw new NotImplementedException();

        public bool GetBoolean(int i) => throw new NotImplementedException();

        public byte GetByte(int i) => throw new NotImplementedException();

        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) => throw new NotImplementedException();

        public CellError? GetCellError(int i) => throw new NotImplementedException();

        public CellStyle GetCellStyle(int i) => throw new NotImplementedException();

        public char GetChar(int i) => throw new NotImplementedException();

        public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) => throw new NotImplementedException();

        public double GetColumnWidth(int i) => throw new NotImplementedException();

        public IDataReader GetData(int i) => throw new NotImplementedException();

        public string GetDataTypeName(int i) => throw new NotImplementedException();

        public DateTime GetDateTime(int i) => throw new NotImplementedException();

        public decimal GetDecimal(int i) => throw new NotImplementedException();

        public double GetDouble(int i) => throw new NotImplementedException();

        [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields | DynamicallyAccessedMemberTypes.PublicProperties)]
        public Type GetFieldType(int i) => throw new NotImplementedException();

        public float GetFloat(int i) => throw new NotImplementedException();

        public Guid GetGuid(int i) => throw new NotImplementedException();

        public short GetInt16(int i) => throw new NotImplementedException();

        public int GetInt32(int i) => throw new NotImplementedException();

        public long GetInt64(int i) => throw new NotImplementedException();

        public string GetName(int i) => throw new NotImplementedException();

        public int GetNumberFormatIndex(int i) => throw new NotImplementedException();

        public string GetNumberFormatString(int i) => throw new NotImplementedException();

        public string GetNumberFormatString(int i, IFormatProvider provider) => throw new NotImplementedException();

        public int GetOrdinal(string name) => throw new NotImplementedException();

        public DataTable? GetSchemaTable() => throw new NotImplementedException();

        public string GetString(int i) => throw new NotImplementedException();

        public Func<int, object>? GetValueAction { get; set; }

        public object GetValue(int i) => GetValueAction != null ? GetValueAction(i) : throw new NotImplementedException();

        public int GetValues(object[] values) => throw new NotImplementedException();

        public bool IsDBNull(int i) => throw new NotImplementedException();

        public bool NextResult() => throw new NotImplementedException();

        public bool Read() => throw new NotImplementedException();

        public void Reset() => throw new NotImplementedException();
    }
}
