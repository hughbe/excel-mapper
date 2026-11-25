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

    // The second column keeps the row from being blank when the cell is empty, as a blank row at
    // the end of a sheet isn't a row.
    private static T Map<T>(object? cell)
    {
        using var importer = Create(["Value", "Other"], [cell, 1]);
        return importer.ReadSheet().ReadRows<Holder<T>>().Single().Value;
    }

    private static ExcelMappingException MapFails<T>(object? cell)
    {
        using var importer = Create(["Value", "Other"], [cell, 1]);
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
    public void Map_FractionalNumberToInteger_ThrowsExcelMappingException()
    {
        // Excel stores numbers as doubles. The fractional part isn't silently dropped, matching
        // text cells where "2.7" can't be parsed as an int, with the same exception.
        var ex = MapFails<int>(2.7);
        Assert.IsType<OverflowException>(ex.InnerException);
        MapFails<long>(-0.5);
        MapFails<byte>(1.25);
#if LIBRARY_NET8
        MapFails<System.Numerics.BigInteger>(3.5);
#endif
        Assert.Null(Map<int?>(null));
    }

    [Fact]
    public void Map_WholeNumberDoubleToInteger_Success()
    {
        Assert.Equal(2, Map<int>(2.0));
        Assert.Equal(-3L, Map<long>(-3.0));
        Assert.Equal(1e15, (double)Map<long>(1e15));
    }

    [Fact]
    public void Map_FractionalNumberToFloatingPoint_Success()
    {
        Assert.Equal(2.7, Map<double>(2.7));
        Assert.Equal(2.7m, Map<decimal>(2.7));
    }

    public class FallbackRow
    {
        [ExcelInvalidValue(-1)]
        public int Value { get; set; }
    }

    [Fact]
    public void Map_FractionalNumberWithInvalidFallback_UsesFallback()
    {
        using var importer = Create(["Value"], [2.7]);
        Assert.Equal(-1, importer.ReadSheet().ReadRows<FallbackRow>().Single().Value);
    }

    [Fact]
    public void Map_NumberToString_UsesInvariantCulture()
    {
        Assert.Equal("42", Map<string>(42));

        // Not the machine's culture.
        var culture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            Assert.Equal("2.5", Map<string>(2.5));

            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
            Assert.Equal("2.5", Map<string>(2.5));
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
#if LIBRARY_NET8
        Assert.Equal(new DateOnly(2024, 3, 4), Map<DateOnly>(new DateOnly(2024, 3, 4)));
#endif
        Assert.Equal(date, Map<DateTimeOffset>(date).DateTime);
        Assert.Equal("03/04/2024 00:00:00", Map<string>(date));
    }

    [Fact]
    public void Map_SerialNumberWithoutDateFormat_IsDate()
    {
        // 45355 is how Excel stores 2024-03-04. Without a date format it is read as a number.
        var serial = new DateTime(2024, 3, 4, 12, 0, 0).ToOADate();
        Assert.Equal(new DateTime(2024, 3, 4, 12, 0, 0), Map<DateTime>(serial));
#if LIBRARY_NET8
        Assert.Equal(new DateOnly(2024, 3, 4), Map<DateOnly>(serial));
#endif
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
#if LIBRARY_NET8
        Assert.Equal(new TimeOnly(5, 6, 7), Map<TimeOnly>(new TimeOnly(5, 6, 7)));
#endif
#if LIBRARY_NET8
        Assert.Equal(new TimeSpan(5, 6, 7), Map<TimeSpan>(new TimeOnly(5, 6, 7)));
#endif
        Assert.Equal(TimeSpan.FromHours(30), Map<TimeSpan>(new Duration(TimeSpan.FromHours(30))));
    }

    [Fact]
    public void Map_NumberWithoutTimeFormat_IsFractionOfDay()
    {
#if LIBRARY_NET8
        Assert.Equal(new TimeOnly(6, 0), Map<TimeOnly>(0.25));
#endif
        Assert.Equal(TimeSpan.FromHours(36), Map<TimeSpan>(1.5));
#if LIBRARY_NET8
        MapFails<TimeOnly>(1.5);
#endif
    }

    // ---------- Formulas and errors ----------

    [Fact]
    public void Map_FormulaCells_UseCachedResult()
    {
        Assert.Equal(3, Map<int>(new Formula("1+2", 3)));
        Assert.Equal("ab", Map<string>(new Formula("\"a\"&\"b\"", "ab")));
    }

    [Theory]
    [InlineData("#N/A")]
    [InlineData("#DIV/0!")]
    [InlineData("#VALUE!")]
    [InlineData("#REF!")]
    [InlineData("#NAME?")]
    [InlineData("#NUM!")]
    [InlineData("#NULL!")]
    public void Map_ErrorCells_ThrowsExcelMappingException(string error)
    {
        // A formula error isn't a value of any type, nor an empty cell.
        var exception = MapFails<int?>(new Error(error));
        Assert.Equal($"The cell contains the error {error}, which can't be assigned to member \"Value\" of type int? in column \"Value\" at cell A2 in sheet \"Sheet1\". Fix the formula, or give the member an invalid fallback, such as [ExcelInvalidValue].", exception.Message);

        MapFails<int>(new Error(error));
        MapFails<string>(new Error(error));
        MapFails<object>(new Error(error));
        MapFails<DateTime?>(new Error(error));
    }

    public class ErrorFallbackRow
    {
        [ExcelInvalidValue(null)]
        public int? Number { get; set; }

        [ExcelInvalidValue("missing")]
        public string? Text { get; set; }

        [ExcelDefaultValue(-1)]
        public int Other { get; set; }
    }

    [Fact]
    public void Map_ErrorCellsWithInvalidFallback_UseFallback()
    {
        using var importer = Create(["Number", "Text", "Other"], [new Error("#N/A"), new Error("#REF!"), 1]);
        var row = importer.ReadSheet().ReadRows<ErrorFallbackRow>().Single();
        Assert.Null(row.Number);
        Assert.Equal("missing", row.Text);
        Assert.Equal(1, row.Other);
    }

    private sealed class ErrorTextFallback : ICellFallback<string>
    {
        public string Fallback(CellValue value, CellContext context, Exception? exception)
            => value.Kind == CellKind.Error ? $"error {value.Text} ({value.Value}, {exception is null})" : "other";
    }

    [Fact]
    public void Map_ErrorCellsWithComputedFallback_SeeError()
    {
        using var importer = Create(["Value", "Other"], [new Error("#DIV/0!"), 1]);
        importer.Configuration.RegisterClassMap<Holder<string>>(c => c.Map(r => r.Value).WithInvalidFallback(new ErrorTextFallback()));
        Assert.Equal("error #DIV/0! (#DIV/0!, True)", importer.ReadSheet().ReadRows<Holder<string>>().Single().Value);
    }

    [Fact]
    public void Map_ErrorCellsWithConverterOrTransformer_NotConverted()
    {
        // Converters and transformers don't see errors: the invalid fallback does.
        using var importer = Create(["Value", "Other"], [new Error("#N/A"), 1]);
        importer.Configuration.RegisterClassMap<Holder<string>>(c => c.Map(r => r.Value).WithTrim().WithConverter(value => "converted").WithInvalidFallback("fallback"));
        Assert.Equal("fallback", importer.ReadSheet().ReadRows<Holder<string>>().Single().Value);
    }

    [Fact]
    public void Map_ErrorCellsPreserveFormatting_ThrowsExcelMappingException()
    {
        using var importer = Create(["Value", "Other"], [new Error("#N/A"), 1]);
        importer.Configuration.RegisterClassMap<Holder<string>>(c => c.Map(r => r.Value).MakePreserveFormatting());
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<Holder<string>>().ToArray());
        Assert.StartsWith("The cell contains the error #N/A", exception.Message);
    }

    [Fact]
    public void Map_ErrorCellsInCollectionsAndDictionaries_ThrowsExcelMappingException()
    {
        using var importer = Create(["A", "B"], [1, new Error("#N/A")]);
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<Dictionary<string, int?>>().ToArray());
        Assert.StartsWith("The cell contains the error #N/A", exception.Message);

        using var split = Create(["Value", "Other"], [new Error("#VALUE!"), 1]);
        exception = Assert.Throws<ExcelMappingException>(() => split.ReadSheet().ReadRows<Holder<string[]>>().ToArray());
        Assert.StartsWith("The cell contains the error #VALUE!", exception.Message);

        using var splitFallback = Create(["Value", "Other"], [new Error("#VALUE!"), 1]);
        splitFallback.Configuration.RegisterClassMap<Holder<string[]>>(c => c.MapCollection<string>(r => r.Value).WithElementMap(e => e.WithInvalidFallback("x")));
        Assert.Equal(["x"], splitFallback.ReadSheet().ReadRows<Holder<string[]>>().Single().Value);
    }

    [Fact]
    public void ReadRowResults_ErrorCells_ReportInvalidValue()
    {
        using var importer = Create(["Value", "Other"], [new Error("#REF!"), 1], [2, 1]);
        var results = importer.ReadSheet().ReadRowResults<Holder<int>>().ToArray();
        var error = Assert.Single(results[0].Errors);
        Assert.StartsWith("The cell contains the error #REF!", error.Message);
        Assert.Equal(2, results[1].Value!.Value);
    }

    [Fact]
    public void ReadRows_RowOfOnlyErrors_IsNotBlank()
    {
        // A row whose cells are all errors has values, so it isn't a blank row at the end of the
        // sheet, and SkipBlankLines doesn't skip it.
        using var importer = Create(["Value"], [1], [new Error("#N/A")]);
        importer.Configuration.SkipBlankLines = true;
        importer.Configuration.RegisterClassMap<Holder<int?>>(c => c.Map(r => r.Value).WithInvalidFallback(-1));
        Assert.Equal([1, -1], importer.ReadSheet().ReadRows<Holder<int?>>().Select(r => r.Value));
    }

    // ---------- object members ----------

    [Fact]
    public void Map_ObjectMembers_ReceiveTypedValues()
    {
        // object members, and object values in dictionaries and untyped collections, receive the
        // value as Excel stores it.
        Assert.Equal(2.5, Map<object>(2.5));
        Assert.Equal(new DateTime(2024, 3, 4), Map<object>(new DateTime(2024, 3, 4)));
        Assert.Equal(true, Map<object>(true));
        Assert.Equal("text", Map<object>("text"));
        Assert.Equal(3.0, Map<object>(new Formula("1+2", 3)));
        Assert.Null(Map<object>(null));
        Assert.IsType<double>(Map<IConvertible>(2.5));

        using var importer = Create(["Number", "Date", "Flag", "Text"], [2.5, new DateTime(2024, 3, 4), false, "x"]);
        var row = importer.ReadSheet().ReadRows<Dictionary<string, object>>().Single();
        Assert.Equal(2.5, row["Number"]);
        Assert.Equal(new DateTime(2024, 3, 4), row["Date"]);
        Assert.Equal(false, row["Flag"]);
        Assert.Equal("x", row["Text"]);
    }

    public class ObjectRow
    {
        public object? Value { get; set; }
        public string? Text { get; set; }
    }

    [Fact]
    public void Map_ObjectMemberWithPreserveFormatting_ReceivesDisplayedText()
    {
        using var importer = Create(["Value"], [new Formatted(0.25, "0.00%")]);
        importer.Configuration.RegisterClassMap<ObjectRow>(c => c.Map(p => p.Value).MakePreserveFormatting());
        Assert.Equal("25.00%", importer.ReadSheet().ReadRows<ObjectRow>().Single().Value);
    }

    [Fact]
    public void Map_StringCollection_ReceivesText()
    {
        // StringCollection only accepts strings, so its elements are still converted to text.
        using var importer = Create(["A", "B"], [1.5, "x"]);
        var row = importer.ReadSheet().ReadRows<System.Collections.Specialized.StringCollection>().Single();
        Assert.Equal(["1.5", "x"], row.Cast<string>());
    }

    [Fact]
    public void Map_CsvObjectMembers_ReceiveText()
    {
        // CSV cells are always text.
        using var importer = Helpers.GetCsvImporter("Value\n2.5\n");
        Assert.Equal("2.5", importer.ReadSheet().ReadRows<Holder<object>>().Single().Value);
    }
}
