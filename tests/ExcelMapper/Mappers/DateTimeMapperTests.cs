using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers.Tests;

public class DateTimeMapperTests
{
    [Fact]
    public void Ctor_Default()
    {
        var item = new DateTimeMapper();
        Assert.Equal(["G"], item.Formats);
        Assert.Null(item.Provider);
        Assert.Equal(DateTimeStyles.None, item.Style);
    }

    [Fact]
    public void Formats_SetValid_GetReturnsExpected()
    {
        var formats = new string[] { "abc" };
        var item = new DateTimeMapper
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
        var item = new DateTimeMapper();
        Assert.Throws<ArgumentNullException>("value", () => item.Formats = null!);
    }

    [Fact]
    public void Formats_SetEmpty_ThrowsArgumentException()
    {
        var item = new DateTimeMapper();
        Assert.Throws<ArgumentException>("value", () => item.Formats = []);
    }

    [Fact]
    public void Formats_SetNullValueInValue_ThrowsArgumentException()
    {
        var item = new DateTimeMapper();
        Assert.Throws<ArgumentException>("value", () => item.Formats = [null!]);
    }

    [Fact]
    public void Formats_SetEmptyValueInValue_ThrowsArgumentException()
    {
        var item = new DateTimeMapper();
        Assert.Throws<ArgumentException>("value", () => item.Formats = [""]);
    }

    [Fact]
    public void Provider_Set_GetReturnsExpected()
    {
        var provider = CultureInfo.CurrentCulture;
        var item = new DateTimeMapper
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
        var item = new DateTimeMapper
        {
            Style = style
        };
        Assert.Equal(style, item.Style);

        // Set same.
        item.Style = style;
        Assert.Equal(style, item.Style);
    }

    public static IEnumerable<object?[]> Map_ValidStringValue_TestData()
    {
        yield return new object?[] { new DateTime(2017, 7, 12, 7, 57, 46).ToString("G"), new string[] { "G" }, null, DateTimeStyles.None, new DateTime(2017, 7, 12, 7, 57, 46) };
        yield return new object?[] { new DateTime(2017, 7, 12, 7, 57, 46).ToString("G"), new string[] { "G", "yyyy-MM-dd" }, null, DateTimeStyles.None, new DateTime(2017, 7, 12, 7, 57, 46) };
        yield return new object?[] { "   2017-07-12   ", new string[] { "G", "yyyy-MM-dd" }, null, DateTimeStyles.AllowWhiteSpaces, new DateTime(2017, 7, 12) };
        yield return new object?[] { new DateTime(2017, 7, 12, 7, 57, 46).ToString("R"), new string[] { "yyyy-MM-dd", "R" }, null, DateTimeStyles.None, new DateTime(2017, 7, 12, 7, 57, 46) };

        var frenchCulture = new CultureInfo("fr-FR");
        yield return new object?[] { new DateTime(2017, 10, 12, 7, 57, 46).ToString("yyyy-MMMM-dd HH:mm:ss", frenchCulture), new string[] { "yyyy-MMMM-dd HH:mm:ss" }, frenchCulture, DateTimeStyles.None, new DateTime(2017, 10, 12, 7, 57, 46) };
    }

    [Theory]
    [MemberData(nameof(Map_ValidStringValue_TestData))]
    public void Map_ValidStringValue_ReturnsSuccess(string stringValue, string[] formats, IFormatProvider? provider, DateTimeStyles style, DateTime expected)
    {
        var item = new DateTimeMapper
        {
            Formats = formats,
            Provider = provider,
            Style = style
        };

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, Assert.IsType<DateTime>(result.Value));
        Assert.Null(result.Exception);
    }

    public static IEnumerable<object?[]> Map_DefaultFormatsStringValue_TestData()
    {
        // ISO 8601.
        yield return new object?[] { "2024-03-04", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4) };
        yield return new object?[] { "2024-03-04", CultureInfo.GetCultureInfo("en-US"), new DateTime(2024, 3, 4) };
        yield return new object?[] { "2024-03-04T10:30", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4, 10, 30, 0) };
        yield return new object?[] { "2024-03-04T10:30:15", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4, 10, 30, 15) };
        yield return new object?[] { "2024-03-04T10:30:15.123", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4, 10, 30, 15, 123) };
        yield return new object?[] { "2024-03-04 10:30", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4, 10, 30, 0) };
        yield return new object?[] { "2024-03-04 10:30:15", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4, 10, 30, 15) };
        // Times with an offset are converted to local time, as DateTime.Parse does.
        yield return new object?[] { "2024-03-04T10:30:15Z", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4, 10, 30, 15, DateTimeKind.Utc).ToLocalTime() };
        yield return new object?[] { "2024-03-04T10:30:15+01:00", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4, 9, 30, 15, DateTimeKind.Utc).ToLocalTime() };

        // The culture's short date and short date and time.
        yield return new object?[] { "04/03/2024", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4) };
        yield return new object?[] { "3/4/2024", CultureInfo.GetCultureInfo("en-US"), new DateTime(2024, 3, 4) };
        yield return new object?[] { "04.03.2024", CultureInfo.GetCultureInfo("de-DE"), new DateTime(2024, 3, 4) };
        yield return new object?[] { "04/03/2024 10:30", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4, 10, 30, 0) };
        yield return new object?[] { "3/4/2024 10:30 AM", CultureInfo.GetCultureInfo("en-US"), new DateTime(2024, 3, 4, 10, 30, 0) };

        // ICU puts U+202F before AM/PM in en-US patterns; text with an ordinary space still matches.
        yield return new object?[] { "3/4/2024 10:30:15 AM", CultureInfo.GetCultureInfo("en-US"), new DateTime(2024, 3, 4, 10, 30, 15) };
        yield return new object?[] { "3/4/2024 10:30:15 PM", CultureInfo.GetCultureInfo("en-US"), new DateTime(2024, 3, 4, 22, 30, 15) };
        yield return new object?[] { "3/4/2024 10:30:15\u202FAM", CultureInfo.GetCultureInfo("en-US"), new DateTime(2024, 3, 4, 10, 30, 15) };

        // The default format still comes first.
        yield return new object?[] { "04/03/2024 10:30:15", CultureInfo.GetCultureInfo("en-GB"), new DateTime(2024, 3, 4, 10, 30, 15) };
    }

    [Theory]
    [MemberData(nameof(Map_DefaultFormatsStringValue_TestData))]
    public void Map_DefaultFormatsStringValue_ReturnsSuccess(string stringValue, IFormatProvider provider, DateTime expected)
    {
        var item = new DateTimeMapper { Provider = provider };
        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, Assert.IsType<DateTime>(result.Value));
        Assert.Null(result.Exception);
    }

    [Theory]
    [InlineData("2024-3-4")]
    [InlineData("2024-13-01")]
    [InlineData("10:30")]
    [InlineData("1/2")]
    [InlineData("March 2024")]
    public void Map_DefaultFormatsInvalidStringValue_ReturnsInvalid(string stringValue)
    {
        var item = new DateTimeMapper { Provider = CultureInfo.GetCultureInfo("en-GB") };
        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.IsType<FormatException>(result.Exception);
    }

    [Fact]
    public void Map_FormatsSetIsoStringValue_ReturnsInvalid()
    {
        // Setting the formats, even to the default, only accepts those formats.
        var item = new DateTimeMapper { Formats = ["G"], Provider = CultureInfo.GetCultureInfo("en-GB") };
        var result = item.Map(new ReadCellResult(0, "2024-03-04", preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.IsType<FormatException>(result.Exception);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("invalid")]
    [InlineData("12/07/2017 07:57:61")]
    public void Map_InvalidStringValue_ReturnsInvalid(string? stringValue)
    {
        var item = new DateTimeMapper();
        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.NotNull(result.Exception);
    }

    public static IEnumerable<object?[]> Map_ValidDateTimeValue_TestData()
    {
        yield return new object?[] { new DateTime(2017, 7, 12, 7, 57, 46), new string[] { "G" }, null, DateTimeStyles.None, new DateTime(2017, 7, 12, 7, 57, 46) };
        yield return new object?[] { new DateTime(2017, 7, 12, 7, 57, 46), new string[] { "G", "yyyy-MM-dd" }, null, DateTimeStyles.None, new DateTime(2017, 7, 12, 7, 57, 46) };
    }

    [Theory]
    [MemberData(nameof(Map_ValidDateTimeValue_TestData))]
    public void Map_ValidDateTimeValue_ReturnsSuccess(DateTime dateTimeValue, string[] formats, IFormatProvider? provider, DateTimeStyles style, DateTime expected)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => dateTimeValue
        };

        var item = new DateTimeMapper
        {
            Formats = formats,
            Provider = provider,
            Style = style
        };

        var result = item.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, Assert.IsType<DateTime>(result.Value));
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_InvalidFormats_ThrowsFormatException()
    {
        var item = new DateTimeMapper
        {
            Formats = ["Invalid"]
        };

        var result = item.Map(new ReadCellResult(0, new DateTime(2020, 1, 1).ToString(), preserveFormatting: false));
        Assert.IsType<FormatException>(result.Exception);
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

    [Fact]
    public void Map_SerialNumberValue_ReturnsSuccess()
    {
        var expected = new DateTime(2024, 3, 4, 12, 0, 0);
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => expected.ToOADate()
        };
        var item = new DateTimeMapper();

        var result = item.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(expected, Assert.IsType<DateTime>(result.Value));
        Assert.Null(result.Exception);
    }

    [Theory]
    [InlineData(1e10)]
    [InlineData(-1e10)]
    public void Map_SerialNumberValueOutOfRange_ReturnsInvalid(double value)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => value
        };
        var item = new DateTimeMapper();

        var result = item.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.NotNull(result.Exception);
    }
}
