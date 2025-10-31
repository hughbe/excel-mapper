using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Numerics;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers.Tests;

public class INumberBaseMapperTests
{
    [Fact]
    public void Ctor_Default()
    {
        var mapper = new INumberBaseMapper<BigInteger>();
        Assert.Equal(NumberStyles.Number, mapper.Style);
        Assert.Null(mapper.Provider);
    }

    [Theory]
    [InlineData(NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite)]
    [InlineData(NumberStyles.HexNumber)]
    [InlineData(NumberStyles.Number)]
    [InlineData(NumberStyles.None)]
    public void Style_Set_GetReturnsExpected(NumberStyles value)
    {
        var mapper = new INumberBaseMapper<BigInteger>
        {
            Style = value
        };
        Assert.Equal(value, mapper.Style);

        // Set same.
        mapper.Style = value;
        Assert.Equal(value, mapper.Style);

        // Set different.
        mapper.Style = NumberStyles.Integer;
        Assert.Equal(NumberStyles.Integer, mapper.Style);
    }

    public static IEnumerable<object?[]> Provider_Set_GetReturnsExpected_Data()
    {
        yield return new object?[] { CultureInfo.CurrentCulture };
        yield return new object?[] { CultureInfo.InvariantCulture };
        yield return new object?[] { null };
    }

    [Theory]
    [MemberData(nameof(Provider_Set_GetReturnsExpected_Data))]
    public void Provider_Set_GetReturnsExpected(IFormatProvider? value)
    {
        var mapper = new INumberBaseMapper<BigInteger>
        {
            Provider = value
        };
        Assert.Same(value, mapper.Provider);

        // Set same.
        mapper.Provider = value;
        Assert.Same(value, mapper.Provider);

        // Set null.
        mapper.Provider = null;
        Assert.Null(mapper.Provider);
    }

    [Theory]
    [InlineData(123.0, 123)]
    [InlineData(0.0, 0)]
    [InlineData(-456.0, -456)]
    [InlineData(999999999999.0, 999999999999)]
    public void Map_ValidDoubleValue_ReturnsSuccess(double doubleValue, long expected)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => doubleValue
        };

        var mapper = new INumberBaseMapper<BigInteger>();
        var result = mapper.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(new BigInteger(expected), Assert.IsType<BigInteger>(result.Value));
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_ValidStringValue_ReturnsSuccess()
    {
        var mapper = new INumberBaseMapper<BigInteger>();
        
        var result1 = mapper.Map(new ReadCellResult(0, "123", preserveFormatting: false));
        Assert.True(result1.Succeeded);
        Assert.Equal(new BigInteger(123), Assert.IsType<BigInteger>(result1.Value));
        Assert.Null(result1.Exception);

        var result2 = mapper.Map(new ReadCellResult(0, "0", preserveFormatting: false));
        Assert.True(result2.Succeeded);
        Assert.Equal(new BigInteger(0), Assert.IsType<BigInteger>(result2.Value));
        Assert.Null(result2.Exception);

        var result3 = mapper.Map(new ReadCellResult(0, "-456", preserveFormatting: false));
        Assert.True(result3.Succeeded);
        Assert.Equal(new BigInteger(-456), Assert.IsType<BigInteger>(result3.Value));
        Assert.Null(result3.Exception);

        // Very large number
        var result4 = mapper.Map(new ReadCellResult(0, "999999999999999999999999", preserveFormatting: false));
        Assert.True(result4.Succeeded);
        Assert.True(Assert.IsType<BigInteger>(result4.Value) > 0);
        Assert.Null(result4.Exception);
    }

    [Theory]
    [InlineData("1,234", NumberStyles.AllowThousands, 1234)]
    [InlineData("(123)", NumberStyles.AllowParentheses, -123)]
    [InlineData(" 123 ", NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite, 123)]
    public void Map_ValidStringValueWithNumberStyle_ReturnsSuccess(string stringValue, NumberStyles style, long expected)
    {
        var mapper = new INumberBaseMapper<BigInteger>
        {
            Style = style
        };
        var result = mapper.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal(new BigInteger(expected), Assert.IsType<BigInteger>(result.Value));
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_ValidStringValueWithProvider_ReturnsSuccess()
    {
        // German culture with period as thousands separator and comma as decimal
        var germanCulture = new CultureInfo("de-DE");
        var mapper1 = new INumberBaseMapper<BigInteger>
        {
            Style = NumberStyles.Integer,
            Provider = germanCulture
        };
        var result1 = mapper1.Map(new ReadCellResult(0, "1234", preserveFormatting: false));
        Assert.True(result1.Succeeded);
        Assert.Equal(new BigInteger(1234), Assert.IsType<BigInteger>(result1.Value));
        Assert.Null(result1.Exception);

        // US culture
        var usCulture = new CultureInfo("en-US");
        var mapper2 = new INumberBaseMapper<BigInteger>
        {
            Style = NumberStyles.Integer,
            Provider = usCulture
        };
        var result2 = mapper2.Map(new ReadCellResult(0, "1234", preserveFormatting: false));
        Assert.True(result2.Succeeded);
        Assert.Equal(new BigInteger(1234), Assert.IsType<BigInteger>(result2.Value));
        Assert.Null(result2.Exception);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("invalid")]
    [InlineData("abc123")]
    public void Map_InvalidStringValue_ReturnsInvalid(string? stringValue)
    {
        var mapper = new INumberBaseMapper<BigInteger>();
        var result = mapper.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.NotNull(result.Exception);
    }

    [Fact]
    public void Map_Int128_ValidDoubleValue_ReturnsSuccess()
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => 123.0
        };

        var mapper = new INumberBaseMapper<Int128>();
        var result = mapper.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal((Int128)123, Assert.IsType<Int128>(result.Value));
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_Int128_ValidStringValue_ReturnsSuccess()
    {
        var mapper = new INumberBaseMapper<Int128>();
        var result = mapper.Map(new ReadCellResult(0, "123", preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal((Int128)123, Assert.IsType<Int128>(result.Value));
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_UInt128_ValidDoubleValue_ReturnsSuccess()
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => 123.0
        };

        var mapper = new INumberBaseMapper<UInt128>();
        var result = mapper.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal((UInt128)123, Assert.IsType<UInt128>(result.Value));
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_UInt128_ValidStringValue_ReturnsSuccess()
    {
        var mapper = new INumberBaseMapper<UInt128>();
        var result = mapper.Map(new ReadCellResult(0, "123", preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal((UInt128)123, Assert.IsType<UInt128>(result.Value));
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_Half_ValidDoubleValue_ReturnsSuccess()
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => 123.5
        };

        var mapper = new INumberBaseMapper<Half>();
        var result = mapper.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal((Half)123.5, Assert.IsType<Half>(result.Value));
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_Half_ValidStringValue_ReturnsSuccess()
    {
        var mapper = new INumberBaseMapper<Half>();
        var result = mapper.Map(new ReadCellResult(0, "123.5", preserveFormatting: false));
        Assert.True(result.Succeeded);
        Assert.Equal((Half)123.5, Assert.IsType<Half>(result.Value));
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_Complex_ValidStringValue_ReturnsSuccess()
    {
        var mapper = new INumberBaseMapper<Complex>();
        // Complex numbers need to be in the format "<real, imaginary>"
        var result = mapper.Map(new ReadCellResult(0, "<3; 4>", preserveFormatting: false));
        Assert.True(result.Succeeded);
        var complexValue = Assert.IsType<Complex>(result.Value);
        Assert.Equal(3, complexValue.Real);
        Assert.Equal(4, complexValue.Imaginary);
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Map_DoubleOverflow_ReturnsInvalid()
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => 999999999.0
        };

        var mapper = new INumberBaseMapper<byte>();
        var result = mapper.Map(new ReadCellResult(0, reader, preserveFormatting: false));
        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.NotNull(result.Exception);
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
