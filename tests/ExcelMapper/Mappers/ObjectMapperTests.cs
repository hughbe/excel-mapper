using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers.Tests;

public class ObjectMapperTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("abc")]
    public void Map_StringValue_ReturnsString(string? stringValue)
    {
        var item = new ObjectMapper();

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(stringValue, result.Value);
        Assert.Null(result.Exception);
        Assert.Equal(CellMapperResult.HandleAction.UseResultAndContinueMapping, result.Action);
    }

    public static IEnumerable<object[]> Map_TypedValue_TestData()
    {
        yield return new object[] { 2.5 };
        yield return new object[] { true };
        yield return new object[] { new DateTime(2024, 3, 4) };
        yield return new object[] { TimeSpan.FromHours(30) };
        yield return new object[] { "text" };
    }

    [Theory]
    [MemberData(nameof(Map_TypedValue_TestData))]
    public void Map_TypedValue_ReturnsValue(object value)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => value
        };
        var item = new ObjectMapper();

        var result = item.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(value, result.Value);
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_EmptyCell_ReturnsNull()
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => null!
        };
        var item = new ObjectMapper();

        var result = item.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Null(result.Value);
    }

    [Fact]
    public void Map_PreserveFormatting_ReturnsString()
    {
        var item = new ObjectMapper();

        var result = item.Map(new ReadCellResult(0, "00123", preserveFormatting: true));
        Assert.True(result.Succeeded);
        Assert.Equal("00123", result.Value);
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
