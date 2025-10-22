using System.Data;
using System.Diagnostics.CodeAnalysis;
using ExcelDataReader;

namespace ExcelMapper.Abstractions.Tests;

public class ReadCellResultTests
{
    [Fact]
    public void Ctor_Default()
    {
        var result = new ReadCellResult();
        Assert.Equal(0, result.ColumnIndex);
        Assert.Null(result.StringValue);
        Assert.Null(result.Reader);
        Assert.False(result.PreserveFormatting);
    }

    [Theory]
    [InlineData(1, null, true)]
    [InlineData(1, null, false)]
    [InlineData(0, "", true)]
    [InlineData(0, "", false)]
    [InlineData(2, "1", true)]
    [InlineData(2, "1", false)]
    [InlineData(2, "stringValue", true)]
    [InlineData(2, "stringValue", false)]
    public void Ctor_ColumnIndex_StringValue_Bool(int columnIndex, string? stringValue, bool preserveFormatting)
    {
        var result = new ReadCellResult(columnIndex, stringValue, preserveFormatting);
        Assert.Equal(columnIndex, result.ColumnIndex);
        Assert.Equal(stringValue, result.StringValue);
        Assert.Null(result.Reader);
        Assert.Equal(preserveFormatting, result.PreserveFormatting);
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(0, false)]
    [InlineData(2, true)]
    [InlineData(2, false)]
    public void Ctor_ColumnIndex_IExcelDataReader_Bool(int columnIndex, bool preserveFormatting)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => "Test"
        };
        if (preserveFormatting)
        {
            reader.GetNumberFormatStringAction = (i) => "General";
        }

        var result = new ReadCellResult(columnIndex, reader, preserveFormatting);
        Assert.Equal(columnIndex, result.ColumnIndex);
        Assert.Equal("Test", result.StringValue);
        Assert.Same(reader, result.Reader);
        Assert.Equal(preserveFormatting, result.PreserveFormatting);
    }

    [Fact]
    public void Ctor_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
    {
        var reader = new MockExcelDataReader();
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => new ReadCellResult(-1, reader, false));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => new ReadCellResult(-1, "stringValue", false));
    }

    [Fact]
    public void Ctor_NullReader_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("reader", () => new ReadCellResult(0, (IExcelDataReader)null!, false));
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData("", true)]
    [InlineData("", false)]
    [InlineData("  ", false)]
    [InlineData("1", true)]
    [InlineData("1", false)]
    [InlineData("stringValue", true)]
    [InlineData("stringValue", false)]
    public void GetString_InvokeStringValue_ReturnExpected(string? stringValue, bool preserveFormatting)
    {
        var result = new ReadCellResult(0, stringValue, preserveFormatting);
        Assert.Equal(stringValue, result.GetString());

        // Call again to test caching.
        Assert.Equal(stringValue, result.GetString());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  ")]
    [InlineData("1")]
    [InlineData("stringValue")]
    public void GetString_InvokeReaderStringValueNoPreserveFormatting_ReturnsExpected(string? stringValue)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => stringValue!
        };
        var result = new ReadCellResult(0, reader, false);
        Assert.Equal(stringValue, result.GetString());

        // Call again to test caching.
        Assert.Equal(stringValue, result.GetString());
    }

    [Fact]
    public void GetString_InvokeReaderNonStringValueNoPreserveFormatting_ReturnsExpected()
    {
        var value = new object();
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => value!
        };
        var result = new ReadCellResult(0, reader, false);
        var stringValue = result.GetString();
        Assert.Equal(value.ToString(), stringValue);

        // Call again to test caching.
        Assert.Same(stringValue, result.GetString());
    }

    [Theory]
    [InlineData(null, null, "")]
    [InlineData(null, "", "")]
    [InlineData(null, "General", "")]
    [InlineData(null, "0.0%", "")]
    [InlineData("", null, "")]
    [InlineData("", "", "")]
    [InlineData("", "General", "")]
    [InlineData("", "0.0%", "")]
    [InlineData(" ", null, " ")]
    [InlineData(" ", "", " ")]
    [InlineData("  ", "General", "  ")]
    [InlineData("  ", "0.0%", "  ")]
    [InlineData("1", null, "1")]
    [InlineData("1", "", "1")]
    [InlineData("1", "General", "1")]
    [InlineData("1", "0.0%", "1")]
    [InlineData("stringValue", null, "stringValue")]
    [InlineData("stringValue", "", "stringValue")]
    [InlineData("stringValue", "General", "stringValue")]
    [InlineData("stringValue", "0.0%", "stringValue")]
    public void GetString_InvokeReaderPreserveFormatting_ReturnsExpected(string? stringValue, string? formatString, string? expected)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => stringValue!,
            GetNumberFormatStringAction = (i) => formatString!
        };
        var result = new ReadCellResult(0, reader, true);
        Assert.Equal(expected, result.GetString());

        // Call again to test caching.
        Assert.Equal(expected, result.GetString());
    }

    [Fact]
    public void GetString_InvokeReaderNonStringValuePreserveFormatting_ReturnsExpected()
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => 42!,
            GetNumberFormatStringAction = (i) => "0.0%"
        };
        var result = new ReadCellResult(0, reader, true);
        var stringValue = result.GetString();
        Assert.Equal("4200.0%", stringValue);

        // Call again to test caching.
        Assert.Same(stringValue, result.GetString());
    }

    [Fact]
    public void GetString_InvokeStringValueMultipleTimes_Caches()
    {
        var stringValue = "stringValue";
        var result = new ReadCellResult(0, stringValue, false);
        Assert.Same(stringValue, result.GetString());

        // Call again to test caching.
        Assert.Same(stringValue, result.GetString());

        // Check GetValue also uses the cache.
        Assert.Same(stringValue, result.GetValue());
    }

    [Fact]
    public void GetString_InvokeReaderNoPreserveFormattingMultipleTimes_Caches()
    {
        var callCount = 0;
        var stringValue = "stringValue";
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) =>
            {
                callCount++;
                return stringValue;
            }
        };
        var result = new ReadCellResult(0, reader, false);
        Assert.Same(stringValue, result.GetString());
        Assert.Equal(1, callCount);

        // Call again to test caching.
        Assert.Same(stringValue, result.GetString());
        Assert.Equal(1, callCount);

        // Check GetValue also uses the cache.
        Assert.Same(stringValue, result.GetValue());
        Assert.Equal(1, callCount);
    }

    [Fact]
    public void GetString_InvokeReaderNonStringNoPreserveFormattingMultipleTimes_Caches()
    {
        var callCount = 0;
        var value = new object();
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) =>
            {
                callCount++;
                return value;
            }
        };
        var result = new ReadCellResult(0, reader, false);
        var stringValue = result.GetString();
        Assert.Equal(stringValue, result.GetString());
        Assert.Equal(1, callCount);

        // Call again to test caching.
        Assert.Same(stringValue, result.GetString());
        Assert.Equal(1, callCount);

        // Check GetValue also uses the cache.
        Assert.Same(value, result.GetValue());
        Assert.Equal(1, callCount);
    }

    [Fact]
    public void GetString_InvokeReaderPreserveFormattingStringMultipleTimes_Caches()
    {
        var getValueCallCount = 0;
        var getNumberFormatStringCallCount = 0;
        var stringValue = "stringValue";
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) =>
            {
                getValueCallCount++;
                return stringValue;
            },
            GetNumberFormatStringAction = (i) =>
            {
                getNumberFormatStringCallCount++;
                return "General";
            }
        };
        var result = new ReadCellResult(0, reader, true);
        Assert.Same(stringValue, result.GetString());
        Assert.Equal(1, getValueCallCount);
        Assert.Equal(1, getNumberFormatStringCallCount);

        // Call again to test caching.
        Assert.Same(stringValue, result.GetString());
        Assert.Equal(1, getValueCallCount);
        Assert.Equal(1, getNumberFormatStringCallCount);

        // Check GetValue also uses the cache.
        Assert.Equal(stringValue, result.GetValue());
        Assert.Equal(1, getValueCallCount);
        Assert.Equal(1, getNumberFormatStringCallCount);
    }

    [Fact]
    public void GetString_InvokeReaderPreserveFormattingIntMultipleTimes_Caches()
    {
        var getValueCallCount = 0;
        var getNumberFormatStringCallCount = 0;
        var intValue = 42;
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) =>
            {
                getValueCallCount++;
                return intValue;
            },
            GetNumberFormatStringAction = (i) =>
            {
                getNumberFormatStringCallCount++;
                return "General";
            }
        };
        var result = new ReadCellResult(0, reader, true);
        var stringValue = result.GetString();
        Assert.Equal("42", stringValue);
        Assert.Equal(1, getValueCallCount);
        Assert.Equal(1, getNumberFormatStringCallCount);

        // Call again to test caching.
        Assert.Same(stringValue, result.GetString());
        Assert.Equal(1, getValueCallCount);
        Assert.Equal(1, getNumberFormatStringCallCount);

        // Check GetValue also uses the cache.
        Assert.Equal(42, result.GetValue());
        Assert.Equal(1, getValueCallCount);
        Assert.Equal(1, getNumberFormatStringCallCount);
    }

    [Fact]
    public void GetString_InvokeDefault_ReturnsNull()
    {
        var result = new ReadCellResult();
        Assert.Null(result.GetString());

        // Call again to test caching.
        Assert.Null(result.GetString());
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData(null, false)]
    [InlineData("", true)]
    [InlineData("", false)]
    [InlineData("  ", true)]
    [InlineData("  ", false)]
    [InlineData("1", true)]
    [InlineData("1", false)]
    [InlineData("stringValue", true)]
    [InlineData("stringValue", false)]
    public void GetValue_InvokeStringValue_ReturnsExpected(string? stringValue, bool preserveFormatting)
    {
        var result = new ReadCellResult(0, stringValue, preserveFormatting);
        Assert.Equal(stringValue, result.GetValue());

        // Call again to test caching.
        Assert.Equal(stringValue, result.GetValue());
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData(null, false)]
    [InlineData("", true)]
    [InlineData("", false)]
    [InlineData("  ", true)]
    [InlineData("  ", false)]
    [InlineData("1", true)]
    [InlineData("1", false)]
    [InlineData("stringValue", true)]
    [InlineData("stringValue", false)]
    public void GetValue_InvokeReaderStringValue_ReturnsExpected(string? stringValue, bool preserveFormatting)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => stringValue!
        };
        var result = new ReadCellResult(0, reader, preserveFormatting);
        Assert.Equal(stringValue, result.GetValue());

        // Call again to test caching.
        Assert.Equal(stringValue, result.GetValue());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void GetValue_InvokeReaderNonStringValue_ReturnsExpected(bool preserveFormatting)
    {
        var value = new object();
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => value
        };
        var result = new ReadCellResult(0, reader, preserveFormatting);
        Assert.Same(value, result.GetValue());

        // Call again to test caching.
        Assert.Same(value, result.GetValue());
    }

    [Fact]
    public void GetValue_InvokeDefault_ReturnsNull()
    {
        var result = new ReadCellResult();
        Assert.Null(result.GetValue());

        // Call again to test caching.
        Assert.Null(result.GetValue());
    }

    [Fact]
    public void GetValue_InvokeStringValueMultipleTimes_Caches()
    {
        var stringValue = "stringValue";
        var result = new ReadCellResult(0, stringValue, false);
        Assert.Same(stringValue, result.GetValue());

        // Call again to test caching.
        Assert.Same(stringValue, result.GetValue());

        // Check GetString also uses the cache.
        Assert.Same(stringValue, result.GetString());
    }

    [Fact]
    public void GetValue_InvokeReaderStringValueNoPreserveFormattingMultipleTimes_Caches()
    {
        var callCount = 0;
        var stringValue = "stringValue";
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) =>
            {
                callCount++;
                return "stringValue";
            }
        };
        var result = new ReadCellResult(0, reader, false);
        Assert.Same(stringValue, result.GetValue());
        Assert.Equal(1, callCount);

        // Call again to test caching.
        Assert.Same(stringValue, result.GetValue());
        Assert.Equal(1, callCount);

        // Check GetString also uses the cache.
        Assert.Same(stringValue, result.GetString());
        Assert.Equal(1, callCount);
    }

    [Fact]
    public void GetValue_InvokeReaderNonStringValueeNoPreserveFormattingMultipleTimes_Caches()
    {
        var callCount = 0;
        var value = new object();
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) =>
            {
                callCount++;
                return value;
            }
        };
        var result = new ReadCellResult(0, reader, false);
        Assert.Same(value, result.GetValue());
        Assert.Equal(1, callCount);

        // Call again to test caching.
        Assert.Same(value, result.GetValue());
        Assert.Equal(1, callCount);

        // Check GetString also uses the cache.
        var stringValue = result.GetString();
        Assert.Equal(value.ToString(), stringValue);
        Assert.Equal(1, callCount);

        // Call again to test caching.
        Assert.Same(stringValue, result.GetString());
        Assert.Equal(1, callCount);
    }

    [Fact]
    public void GetValue_InvokeReaderNonStringValuePreserveFormattingMultipleTimes_Caches()
    {
        var callCount = 0;
        var getNumberFormatStringCallCount = 0;
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) =>
            {
                callCount++;
                return 42;
            },
            GetNumberFormatStringAction = (i) =>
            {
                getNumberFormatStringCallCount++;
                return "0.0%";
            }
        };
        var result = new ReadCellResult(0, reader, true);
        Assert.Equal(42, result.GetValue());
        Assert.Equal(1, callCount);
        Assert.Equal(0, getNumberFormatStringCallCount);

        // Call again to test caching.
        Assert.Equal(42, result.GetValue());
        Assert.Equal(1, callCount);
        Assert.Equal(0, getNumberFormatStringCallCount);

        // Check GetString also uses the cache.
        var stringValue = result.GetString();
        Assert.Equal("4200.0%", stringValue);
        Assert.Equal(1, callCount);
        Assert.Equal(1, getNumberFormatStringCallCount);

        // Call again to test caching.
        Assert.Same(stringValue, result.GetString());
        Assert.Equal(1, callCount);
        Assert.Equal(1, getNumberFormatStringCallCount);
    }

    [Theory]
    [InlineData(null, true, true)]
    [InlineData(null, false, true)]
    [InlineData("", true, true)]
    [InlineData("", false, true)]
    [InlineData(" ", true, false)]
    [InlineData("  ", false, false)]
    [InlineData("1", true, false)]
    [InlineData("1", false, false)]
    [InlineData("stringValue", true, false)]
    [InlineData("stringValue", false, false)]
    public void IsEmpty_InvokeStringValue_ReturnsExpected(string? stringValue, bool preserveFormatting, bool expected)
    {
        var result = new ReadCellResult(0, stringValue, preserveFormatting);
        Assert.Equal(expected, result.IsEmpty());

        // Call again to test caching.
        Assert.Equal(expected, result.IsEmpty());
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData("", true)]
    [InlineData("  ", false)]
    [InlineData("1", false)]
    [InlineData("stringValue", false)]
    public void IsEmpty_InvokeReaderNoPreserveFormatting_ReturnsExpected(string? stringValue, bool expected)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => stringValue!
        };
        var result = new ReadCellResult(0, reader, false);
        Assert.Equal(expected, result.IsEmpty());

        // Call again to test caching.
        Assert.Equal(expected, result.IsEmpty());
    }

    [Theory]
    [InlineData(null, null, true)]
    [InlineData(null, "", true)]
    [InlineData(null, "General", true)]
    [InlineData(null, "0.0%", true)]
    [InlineData("", null, true)]
    [InlineData("", "", true)]
    [InlineData("", "General", true)]
    [InlineData("", "0.0%", true)]
    [InlineData(" ", null, false)]
    [InlineData(" ", "", false)]
    [InlineData("  ", "General", false)]
    [InlineData("  ", "0.0%", false)]
    [InlineData("1", null, false)]
    [InlineData("1", "", false)]
    [InlineData("1", "General", false)]
    [InlineData("1", "0.0%", false)]
    [InlineData("stringValue", null, false)]
    [InlineData("stringValue", "", false)]
    [InlineData("stringValue", "General", false)]
    [InlineData("stringValue", "0.0%", false)]
    public void IsEmpty_InvokeReaderPreserveFormatting_ReturnsExpected(string? stringValue, string? formatString, bool expected)
    {
        var reader = new MockExcelDataReader
        {
            GetValueAction = (i) => stringValue!,
            GetNumberFormatStringAction = (i) => formatString!
        };
        var result = new ReadCellResult(0, reader, true);
        Assert.Equal(expected, result.IsEmpty());

        // Call again to test caching.
        Assert.Equal(expected, result.IsEmpty());
    }

    [Fact]
    public void IsEmpty_InvokeDefault_ReturnsTrue()
    {
        var result = new ReadCellResult();
        Assert.True(result.IsEmpty());

        // Call again to test caching.
        Assert.True(result.IsEmpty());
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

        public Func<int, string> GetNumberFormatStringAction { get; set; } = null!;

        public string GetNumberFormatString(int i) => GetNumberFormatStringAction != null ? GetNumberFormatStringAction(i) : throw new NotImplementedException();

        public int GetOrdinal(string name) => throw new NotImplementedException();

        public DataTable? GetSchemaTable() => throw new NotImplementedException();

        public string GetString(int i) => throw new NotImplementedException();

        public Func<int, object> GetValueAction { get; set; } = null!;

        public object GetValue(int i) => GetValueAction != null ? GetValueAction(i) : throw new NotImplementedException();

        public int GetValues(object[] values) => throw new NotImplementedException();

        public bool IsDBNull(int i) => throw new NotImplementedException();

        public bool NextResult() => throw new NotImplementedException();

        public bool Read() => throw new NotImplementedException();

        public void Reset() => throw new NotImplementedException();
    }
}
