using System.Globalization;
using System.Linq;
using static ExcelMapper.Tests.XlsxBuilder;

namespace ExcelMapper.Tests;

/// <summary>
/// Mapping typed Excel cells (numbers, booleans, dates, times, durations, formulas, errors and
/// formatted values), using workbooks built in memory by <see cref="XlsxBuilder"/>.
/// </summary>
public class TypedCellTests
{
    public class Holder<T>
    {
        public T Value { get; set; } = default!;
    }

    private static T Map<T>(object? cell)
    {
        using var importer = Create(["Value"], [cell]);
        return importer.ReadSheet().ReadRows<Holder<T>>().Single().Value;
    }

    private static ExcelMappingException MapFails<T>(object? cell)
    {
        using var importer = Create(["Value"], [cell]);
        var sheet = importer.ReadSheet();
        return Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Holder<T>>().ToArray());
    }

    // ---------- Numbers ----------

    [Fact]
    public void Map_NumberCells_Success()
    {
        Assert.Equal(42, Map<int>(42));
        Assert.Equal(12345678901L, Map<long>(12345678901L));
        Assert.Equal((byte)200, Map<byte>(200));
        Assert.Equal(1.5f, Map<float>(1.5));
        Assert.Equal(-1.5e10, Map<double>(-1.5e10));
        Assert.Equal(0.1m, Map<decimal>(0.1));
        Assert.Null(Map<int?>(null));
        Assert.Equal(DayOfWeek.Monday, Map<DayOfWeek>(1));
    }

    [Fact]
    public void Map_NumberOutOfRange_ThrowsExcelMappingException()
    {
        var ex = MapFails<byte>(300);
        Assert.Equal("A2", ex.CellAddress);
    }

    [Fact]
    public void Map_NumberToGuid_ThrowsExcelMappingException()
    {
        MapFails<Guid>(1);
    }

    [Fact]
    public void Map_FractionalNumberToInt_Truncates()
    {
        // Documented in the README: Excel stores numbers as doubles, and .NET 8 truncates them.
        Assert.Equal(2, Map<int>(2.7));
    }

    [Fact]
    public void Map_NumberToString_UsesCurrentCulture()
    {
        Assert.Equal("42", Map<string>(42));

        var culture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            Assert.Equal("2.5", Map<string>(2.5));

            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
            Assert.Equal("2,5", Map<string>(2.5));
        }
        finally
        {
            CultureInfo.CurrentCulture = culture;
        }
    }

    [Fact]
    public void Map_FormattedNumbersWithPreserveFormatting_ReturnsDisplayedText()
    {
        using var importer = Create(["Zip", "Rate"], [new Formatted(123, "00000"), new Formatted(0.25, "0.00%")]);
        importer.Configuration.RegisterClassMap<FormattedRow>(c =>
        {
            c.Map(p => p.Zip).MakePreserveFormatting();
            c.Map(p => p.Rate).MakePreserveFormatting();
        });

        var row = importer.ReadSheet().ReadRows<FormattedRow>().Single();
        Assert.Equal("00123", row.Zip);
        Assert.Equal("25.00%", row.Rate);
    }

    public class FormattedRow
    {
        public string Zip { get; set; } = "";
        public string Rate { get; set; } = "";
    }

    [Fact]
    public void Map_PercentToDouble_ReturnsUnderlyingValue()
    {
        Assert.Equal(0.25, Map<double>(new Formatted(0.25, "0.00%")));
    }

    // ---------- Booleans ----------

    [Theory]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(1, true)]
    [InlineData(0, false)]
    [InlineData("TRUE", true)]
    [InlineData("false", false)]
    [InlineData("Yes", true)]
    [InlineData("YES", true)]
    [InlineData("no", false)]
    public void Map_Booleans_Success(object cell, bool expected)
    {
        Assert.Equal(expected, Map<bool>(cell));
    }

    [Theory]
    [InlineData("maybe")]
    [InlineData("y")]
    [InlineData(2)]
    public void Map_InvalidBoolean_ThrowsExcelMappingException(object cell)
    {
        MapFails<bool>(cell);
    }

    [Fact]
    public void Map_BooleanToString_Success()
    {
        Assert.Equal("True", Map<string>(true));
    }

    // ---------- Dates and times ----------

    [Fact]
    public void Map_DateCells_Success()
    {
        var date = new DateTime(2024, 3, 4);
        Assert.Equal(date, Map<DateTime>(date));
        Assert.Equal(new DateTime(2024, 3, 4, 5, 6, 7), Map<DateTime>(new DateTime(2024, 3, 4, 5, 6, 7)));
        Assert.Equal(new DateOnly(2024, 3, 4), Map<DateOnly>(new DateOnly(2024, 3, 4)));
        Assert.Equal(date, Map<DateTimeOffset>(date).DateTime);
        Assert.Equal(date.ToString(), Map<string>(date));
    }

    [Fact]
    public void Map_SerialNumberWithoutDateFormat_IsDate()
    {
        // 45355 is how Excel stores 2024-03-04. Without a date format it is read as a number.
        var serial = new DateTime(2024, 3, 4, 12, 0, 0).ToOADate();
        Assert.Equal(new DateTime(2024, 3, 4, 12, 0, 0), Map<DateTime>(serial));
        Assert.Equal(new DateOnly(2024, 3, 4), Map<DateOnly>(serial));
        Assert.Equal(new DateTime(2024, 3, 4, 12, 0, 0), Map<DateTimeOffset>(serial).DateTime);
        Assert.Equal(new DateTime(2024, 3, 4), Map<DateTime?>(45355));
    }

    [Fact]
    public void Map_SerialNumberOutOfRange_ThrowsExcelMappingException()
    {
        MapFails<DateTime>(1e10);
        MapFails<DateOnly>(-1e10);
    }

    [Fact]
    public void Map_TimeCells_Success()
    {
        Assert.Equal(new TimeOnly(5, 6, 7), Map<TimeOnly>(new TimeOnly(5, 6, 7)));
        Assert.Equal(new TimeSpan(5, 6, 7), Map<TimeSpan>(new TimeOnly(5, 6, 7)));
        Assert.Equal(TimeSpan.FromHours(30), Map<TimeSpan>(new Duration(TimeSpan.FromHours(30))));
    }

    [Fact]
    public void Map_NumberWithoutTimeFormat_IsFractionOfDay()
    {
        Assert.Equal(new TimeOnly(6, 0), Map<TimeOnly>(0.25));
        Assert.Equal(TimeSpan.FromHours(36), Map<TimeSpan>(1.5));
        MapFails<TimeOnly>(1.5);
    }

    // ---------- Formulas and errors ----------

    [Fact]
    public void Map_FormulaCells_UseCachedResult()
    {
        Assert.Equal(3, Map<int>(new Formula("1+2", 3)));
        Assert.Equal("ab", Map<string>(new Formula("\"a\"&\"b\"", "ab")));
    }

    [Fact]
    public void Map_ErrorCells_AreEmpty()
    {
        Assert.Null(Map<int?>(new Error("#N/A")));
        MapFails<int>(new Error("#DIV/0!"));
    }

    // ---------- object members ----------

    [Fact]
    public void Map_ObjectMembers_ReceiveText()
    {
        // object members, and object values in dictionaries, receive the cell's text rather than
        // its typed value. Existing code relies on this, so it is pinned here.
        Assert.Equal("2.5", Map<object>(2.5));

        using var importer = Create(["Number"], [2.5]);
        Assert.Equal("2.5", importer.ReadSheet().ReadRows<Dictionary<string, object>>().Single()["Number"]);
    }
}
