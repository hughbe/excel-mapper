namespace ExcelMapper.Tests;

public class ExcelRangeTests
{
    [Fact]
    public void Ctor_Default()
    {
        var range = new ExcelRange();
        Assert.Equal(Range.All, range.Rows);
        Assert.Equal(Range.All, range.Columns);
    }

    [Theory]
    [InlineData(0, 0, 0, 0)]
    [InlineData(1, 10, 1, 5)]
    [InlineData(5, 15, 3, 8)]
    [InlineData(1, 1, 1, 1)]
    public void Ctor_Int_Int_Int_Int(int rowStart, int rowEnd, int columnStart, int columnEnd)
    {
        var range = new ExcelRange(rowStart, rowEnd, columnStart, columnEnd);
        Assert.Equal(rowStart..(rowEnd + 1), range.Rows);
        Assert.Equal(columnStart..(columnEnd + 1), range.Columns);
        Assert.Equal(rowStart, range.Rows.Start.Value);
        Assert.Equal(rowEnd + 1, range.Rows.End.Value);
        Assert.Equal(columnStart, range.Columns.Start.Value);
        Assert.Equal(columnEnd + 1, range.Columns.End.Value);
    }

    [Fact]
    public void Ctor_NegativeRowStart_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("rowStart", () => new ExcelRange(-1, 10, 1, 5));
    }

    [Fact]
    public void Ctor_NegativeRowEnd_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("rowEnd", () => new ExcelRange(1, -1, 1, 5));
    }

    [Fact]
    public void Ctor_NegativeColumnStart_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("columnStart", () => new ExcelRange(1, 10, -1, 5));
    }

    [Fact]
    public void Ctor_NegativeColumnEnd_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("columnEnd", () => new ExcelRange(1, 10, 1, -1));
    }

    [Fact]
    public void Ctor_RowStartGreaterThanRowEnd_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("rowStart", () => new ExcelRange(6, 5, 1, 5));
    }

    [Fact]
    public void Ctor_ColumnStartGreaterThanColumnEnd_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("columnStart", () => new ExcelRange(1, 10, 6, 5));
    }

    public static IEnumerable<object[]> Range_Range_Data()
    {
        yield return new object[] { 0..5, 0..10 };
        yield return new object[] { 2.., 3..8 };
        yield return new object[] { ..4, ..6 };
        yield return new object[] { .., .. };
    }

    [Theory]
    [MemberData(nameof(Range_Range_Data))]
    public void Ctor_Range_Range(Range rows, Range columns)
    {
        var range = new ExcelRange(rows, columns);
        Assert.Equal(rows, range.Rows);
        Assert.Equal(columns, range.Columns);
    }

    public static IEnumerable<object[]> ValidAddress_Data()
    {
        yield return new object[] { "A1", 0..1, 0..1 };
        yield return new object[] { "AB2", 1..2, 27..28 };
        yield return new object[] { " A1 ", 0..1, 0..1 };
        yield return new object[] { "B2:C5", 1..5, 1..3 };
        yield return new object[] { " B2:C5 ", 1..5, 1..3 };
        yield return new object[] { "AB2:AC3", 1..3, 27..29 };
        yield return new object[] { "D:D", Range.All, 3..4 };
        yield return new object[] { " D:D ", Range.All, 3..4 };
        yield return new object[] { "D:E", Range.All, 3..5 };
        yield return new object[] { " D:E ", Range.All, 3..5 };
        yield return new object[] { "AB:AC", Range.All, 27..29 };
        yield return new object[] { "3:3", 2..3, Range.All };
        yield return new object[] { " 3:3 ", 2..3, Range.All };
        yield return new object[] { "3:7", 2..7, Range.All };
        yield return new object[] { " 3:7 ", 2..7, Range.All };
    }

    [Theory]
    [MemberData(nameof(ValidAddress_Data))]
    public void Ctor_String(string address, Range expectedRows, Range expectedColumns)
    {
        var range = new ExcelRange(address);
        Assert.Equal(expectedRows, range.Rows);
        Assert.Equal(expectedColumns, range.Columns);
    }

    [Fact]
    public void Ctor_NullAddress_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("address", () => new ExcelRange(null!));
    }


    public static IEnumerable<object[]> InvalidAddress_Data()
    {
        // Empty or whitespace.
        yield return new object[] { "" };
        yield return new object[] { " " };

        // Invalid number.
        yield return new object[] { "A0:B10" };
        yield return new object[] { "A1:B0" };
        yield return new object[] { "B1:A10" };
        yield return new object[] { "A1:A0" };
        yield return new object[] { "0:10" };
        yield return new object[] { "1:0" };

        // Not a cell reference.
        yield return new object[] { ":" };
        yield return new object[] { "_" };
        yield return new object[] { "AB" };
        yield return new object[] { "A" };
        yield return new object[] { "A_" };
        yield return new object[] { "invalid" };

        // Invalid first part.
        yield return new object[] { ":A1" };
        yield return new object[] { ":A" };
        yield return new object[] { ":1" };
        yield return new object[] { "_:A1" };
        yield return new object[] { "_:A" };
        yield return new object[] { "_:1" };

        // Invalid second part.
        yield return new object[] { "A1:" };
        yield return new object[] { "A1:_" };
        yield return new object[] { "A1:A_" };
        yield return new object[] { "A1:_1" };
        yield return new object[] { "A1:B" };
        yield return new object[] { "A1:2" };
        yield return new object[] { "A:" };
        yield return new object[] { "A:_" };
        yield return new object[] { "A:A_" };
        yield return new object[] { "A:_A" };
        yield return new object[] { "A:_1" };
        yield return new object[] { "A:1_" };
        yield return new object[] { "A:A1" };
        yield return new object[] { "A:2" };
        yield return new object[] { "1:" };
        yield return new object[] { "1:_" };
        yield return new object[] { "1:A_" };
        yield return new object[] { "1:_1" };
        yield return new object[] { "1:1_" };
        yield return new object[] { "1:A1" };
        yield return new object[] { "1:A" };

        // Invalid range.
        yield return new object[] { "B2:A5" };
        yield return new object[] { "B2:C1" };
        yield return new object[] { "B2:C0" };
        yield return new object[] { "B0:C5" };
        yield return new object[] { "5:3" };
        yield return new object[] { "0:3" };
        yield return new object[] { "3:0" };
        yield return new object[] { "E:D" };

        // Row overflow.
        yield return new object[] { $"A1:B{(long)int.MaxValue + 1}" };
        yield return new object[] { $"A{(long)int.MaxValue + 1}:A2" };
        yield return new object[] { $"1:{(long)int.MaxValue + 1}" };
        yield return new object[] { $"{(long)int.MaxValue + 1}:2" };

        // Multiple colons.
        yield return new object[] { "A1:B2:" };
        yield return new object[] { "A2:B2:C2" };
        yield return new object[] { "A:B:" };
        yield return new object[] { "A:B:C" };
        yield return new object[] { "1:2:" };
        yield return new object[] { "1:2:C" };
    }

    [Theory]
    [MemberData(nameof(InvalidAddress_Data))]
    public void Ctor_InvalidAddress_ThrowsArgumentException(string address)
    {
        Assert.Throws<ArgumentException>("address", () => new ExcelRange(address));
    }

    [Theory]
    [MemberData(nameof(ValidAddress_Data))]
    public void Parse_ValidAddress_ReturnsExpectedRange(string address, Range expectedRows, Range expectedColumns)
    {
        var range = ExcelRange.Parse(address);
        Assert.Equal(expectedRows, range.Rows);
        Assert.Equal(expectedColumns, range.Columns);
    }

    [Fact]
    public void Parse_NullAddress_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("address", () => ExcelRange.Parse(null!));
    }

    [Theory]
    [MemberData(nameof(InvalidAddress_Data))]
    public void Parse_InvalidAddress_ThrowsArgumentException(string address)
    {
        Assert.Throws<ArgumentException>("address", () => ExcelRange.Parse(address));
    }

    [Theory]
    [MemberData(nameof(ValidAddress_Data))]
    public void TryParse_Invoke_ReturnsExpected(string address, Range expectedRows, Range expectedColumns)
    {
        Assert.True(ExcelRange.TryParse(address, out var range));
        Assert.Equal(expectedRows, range.Rows);
        Assert.Equal(expectedColumns, range.Columns);
    }

    [Theory]
    [InlineData(null)]
    [MemberData(nameof(InvalidAddress_Data))]
    public void TryParse_InvalidAddress_ReturnsFalse(string? address)
    {
        Assert.False(ExcelRange.TryParse(address!, out var range));
        Assert.Equal(Range.All, range.Rows);
        Assert.Equal(Range.All, range.Columns);
    }

    [Fact]
    public void Default_IsAllRowsAndColumns()
    {
        ExcelRange range = default;
        Assert.Equal(Range.All, range.Rows);
        Assert.Equal(Range.All, range.Columns);
        Assert.Equal(new ExcelRange(), range);
    }

    public static IEnumerable<object[]> AbsoluteAndLimitAddress_Data()
    {
        yield return new object[] { "$A$1", 0..1, 0..1 };
        yield return new object[] { "$B5", 4..5, 1..2 };
        yield return new object[] { "B$5", 4..5, 1..2 };
        yield return new object[] { "$A$1:$C$10", 0..10, 0..3 };
        yield return new object[] { "$A:$C", Range.All, 0..3 };
        yield return new object[] { "$2:$5", 1..5, Range.All };
        yield return new object[] { "xfd1048576", 1048575..1048576, 16383..16384 };
        yield return new object[] { "A:XFD", Range.All, 0..16384 };
        yield return new object[] { "1:1048576", 0..1048576, Range.All };
    }

    [Theory]
    [MemberData(nameof(AbsoluteAndLimitAddress_Data))]
    public void Parse_AbsoluteAndLimitAddress_ReturnsExpected(string address, Range expectedRows, Range expectedColumns)
    {
        var range = ExcelRange.Parse(address);
        Assert.Equal(expectedRows, range.Rows);
        Assert.Equal(expectedColumns, range.Columns);

        Assert.True(ExcelRange.TryParse(address, out var tryRange));
        Assert.Equal(range, tryRange);
    }

    [Theory]
    // Non-ASCII letters.
    [InlineData("\u00C91")]
    [InlineData("A\u00C91")]
    [InlineData("\u00C9:F")]
    // Beyond the limits of an Excel worksheet.
    [InlineData("XFE1")]
    [InlineData("AAAAAAAAAAAAAAAAAAAA1")]
    [InlineData("A1048577")]
    [InlineData("A:XFE")]
    [InlineData("1:1048577")]
    // Signs and misplaced absolute markers.
    [InlineData("A+1")]
    [InlineData("1:+5")]
    [InlineData("$")]
    [InlineData("$$A1")]
    [InlineData("A$$1")]
    [InlineData("A1$")]
    public void Parse_InvalidExtendedAddress_ThrowsArgumentException(string address)
    {
        Assert.Throws<ArgumentException>("address", () => ExcelRange.Parse(address));
        Assert.False(ExcelRange.TryParse(address, out var range));
        Assert.Equal(new ExcelRange(), range);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("B2:C5")]
    [InlineData("AB2:AC3")]
    [InlineData("D:D")]
    [InlineData("A:XFD")]
    [InlineData("3:7")]
    [InlineData("XFD1048576")]
    public void ToString_Address_RoundTrips(string address)
    {
        var range = ExcelRange.Parse(address);
        Assert.Equal(address, range.ToString());
        Assert.Equal(range, ExcelRange.Parse(range.ToString()));
    }

    [Fact]
    public void ToString_NotAnAddress_DescribesRanges()
    {
        Assert.Equal($"Rows: {Range.All}, Columns: {Range.All}", new ExcelRange().ToString());
        Assert.Equal($"Rows: {2..}, Columns: {1..3}", new ExcelRange(2.., 1..3).ToString());
        Assert.Equal($"Rows: {0..0}, Columns: {0..1}", new ExcelRange(0..0, 0..1).ToString());
    }

    [Fact]
    public void Equals_Invoke_ReturnsExpected()
    {
        var range = new ExcelRange(1..5, 2..3);
        var same = ExcelRange.Parse("C2:C5");
        var different = new ExcelRange(1..5, 2..4);

        Assert.True(range.Equals(same));
        Assert.True(range.Equals((object)same));
        Assert.True(range == same);
        Assert.False(range != same);
        Assert.Equal(range.GetHashCode(), same.GetHashCode());

        Assert.False(range.Equals(different));
        Assert.False(range == different);
        Assert.True(range != different);
        Assert.False(range.Equals(null));
        Assert.False(range.Equals("C2:C5"));
    }
}
