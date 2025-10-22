using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers.Tests;

public class TimeOnlyMapperTests
{
    [Fact]
    public void Ctor_Default()
    {
        var item = new TimeOnlyMapper();
        Assert.Equal(["t"], item.Formats);
        Assert.Null(item.Provider);
        Assert.Equal(DateTimeStyles.None, item.Style);
    }

    [Fact]
    public void Formats_SetValid_GetReturnsExpected()
    {
        var formats = new string[] { "abc" };
        var item = new TimeOnlyMapper
        {
            Formats = formats
        };
        Assert.Same(formats, item.Formats);

        // Set same.
        item.Formats = formats;
        Assert.Same(formats, item.Formats);
    }

    [Fact]
    public void Formats_SetNull_ThrowsArgumentNullException()
    {
        var item = new TimeOnlyMapper();
        Assert.Throws<ArgumentNullException>("value", () => item.Formats = null!);
    }

    [Fact]
    public void Formats_SetEmpty_ThrowsArgumentException()
    {
        var item = new TimeOnlyMapper();
        Assert.Throws<ArgumentException>("value", () => item.Formats = []);
    }

    [Fact]
    public void Formats_SetNullValueInValue_ThrowsArgumentException()
    {
        var item = new TimeOnlyMapper();
        Assert.Throws<ArgumentException>("value", () => item.Formats = [null!]);
    }

    [Fact]
    public void Formats_SetEmptyValueInValue_ThrowsArgumentException()
    {
        var item = new TimeOnlyMapper();
        Assert.Throws<ArgumentException>("value", () => item.Formats = [""]);
    }

    [Fact]
    public void Provider_Set_GetReturnsExpected()
    {
        var provider = CultureInfo.CurrentCulture;
        var item = new TimeOnlyMapper
        {
            Provider = provider
        };
        Assert.Same(provider, item.Provider);

        // Set same.
        item.Provider = provider;
        Assert.Same(provider, item.Provider);

        // Set null.
        item.Provider = null;
        Assert.Null(item.Provider);
    }

    [Theory]
    [InlineData(DateTimeStyles.AdjustToUniversal)]
    [InlineData((DateTimeStyles)int.MaxValue)]
    public void Styles_Set_GetReturnsExpected(DateTimeStyles style)
    {
        var item = new TimeOnlyMapper
        {
            Style = style
        };
        Assert.Equal(style, item.Style);

        // Set same.
        item.Style = style;
        Assert.Equal(style, item.Style);
    }

    public static IEnumerable<object[]> Map_ValidStringValue_TestData()
    {
        yield return new object[] { new TimeOnly(7, 57, 46).ToString("T"), new string[] { "T" }, DateTimeStyles.None, new TimeOnly(7, 57, 46) };
        yield return new object[] { new TimeOnly(7, 57, 46).ToString("T"), new string[] { "T", "yyyy-MM-dd" }, DateTimeStyles.None, new TimeOnly(7, 57, 46) };
        yield return new object[] { $" {new TimeOnly(7, 57, 46):T}  ", new string[] { "yyyy-MM-dd", "T" }, DateTimeStyles.AllowWhiteSpaces, new TimeOnly(7, 57, 46) };
        yield return new object[] { $" {new TimeOnly(7, 57, 46):R}  ", new string[] { "yyyy-MM-dd", "R" }, DateTimeStyles.AllowWhiteSpaces, new TimeOnly(7, 57, 46) };
    }

    [Theory]
    [MemberData(nameof(Map_ValidStringValue_TestData))]
    public void Map_ValidStringValue_ReturnsSuccess(string stringValue, string[] formats, DateTimeStyles style, TimeOnly expected)
    {
        var item = new TimeOnlyMapper
        {
            Formats = formats,
            Style = style
        };

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, Assert.IsType<TimeOnly>(result.Value));
        Assert.Null(result.Exception);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("invalid")]
    [InlineData("12/07/2017 07:57:61")]
    public void Map_InvalidStringValue_ReturnsInvalid(string? stringValue)
    {
        var item = new TimeOnlyMapper();
        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.NotNull(result.Exception);
    }


    public static IEnumerable<object[]> Map_ValidTimeSpan_TestData()
    {
        yield return new object[] { new TimeSpan(7, 57, 46), new string[] { "T" }, DateTimeStyles.None, new TimeOnly(7, 57, 46) };
        yield return new object[] { new TimeSpan(0, 0, 0), new string[] { "T" }, DateTimeStyles.None, new TimeOnly(0, 0, 0) };
        yield return new object[] { new TimeSpan(23, 59, 59), new string[] { "T" }, DateTimeStyles.None, new TimeOnly(23, 59, 59) };
    }

    [Theory]
    [MemberData(nameof(Map_ValidTimeSpan_TestData))]
    public void Map_ValidTimeSpan_ReturnsSuccess(TimeSpan timeSpanValue, string[] formats, DateTimeStyles style, TimeOnly expected)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => timeSpanValue
        };

        var item = new TimeOnlyMapper
        {
            Formats = formats,
            Style = style
        };

        var result = item.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, Assert.IsType<TimeOnly>(result.Value));
        Assert.Null(result.Exception);
    }

    public static IEnumerable<object[]> Map_InvalidTimeSpan_TestData()
    {
        yield return new object[] { -1 };
        yield return new object[] { TimeOnly.MaxValue.Ticks + 1 };
    }

    [Theory]
    [MemberData(nameof(Map_InvalidTimeSpan_TestData))]
    public void Map_InvalidTimeSpan_ReturnsInvalid(long ticks)
    {
        var timeSpan = new TimeSpan(ticks);
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => timeSpan
        };

        var item = new TimeOnlyMapper();
        var result = item.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.IsType<ArgumentOutOfRangeException>(result.Exception);
    }

    [Fact]
    public void Map_InvalidFormats_ThrowsFormatException()
    {
        var item = new TimeOnlyMapper
        {
            Formats = ["Invalid"]
        };

        var result = item.Map(new ReadCellResult(0, new TimeOnly(7, 57, 46).ToString(), preserveFormatting: false));
        Assert.IsType<FormatException>(result.Exception);
    }

    [Fact]
    public void Map_InvalidStyle_ThrowsArgumentException()
    {
        var item = new TimeOnlyMapper
        {
            Style = DateTimeStyles.AssumeLocal | DateTimeStyles.AssumeUniversal
        };

        var result = item.Map(new ReadCellResult(0, new TimeOnly(7, 57, 46).ToString(), preserveFormatting: false));
        Assert.IsType<ArgumentException>(result.Exception);
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
