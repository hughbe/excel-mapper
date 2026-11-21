using System.Globalization;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// The exception for an invalid value is the same type the type's Parse method throws, even where
/// it's created without parsing again.
/// </summary>
public class InvalidValueExceptionTests
{
    public static IEnumerable<object[]> NumberTexts() =>
        new[] { "x", "1.5", "300", "-1", "99999999999999999999999", "1e5", "1e400", "-1e400", "NaN", "Infinity", "-Infinity", "5 x", "$5", "(5)", "--1", "0x1F", "1.0" }
            .Select(t => new object[] { t });

    [Theory]
    [MemberData(nameof(NumberTexts))]
    public void ReadRows_InvalidNumber_ThrowsParseExceptionType(string text)
    {
        AssertSameException<sbyte>(text, s => sbyte.Parse(s, NumberStyles.Number, CultureInfo.CurrentCulture));
        AssertSameException<byte>(text, s => byte.Parse(s, NumberStyles.Number, CultureInfo.CurrentCulture));
        AssertSameException<short>(text, s => short.Parse(s, NumberStyles.Number, CultureInfo.CurrentCulture));
        AssertSameException<ushort>(text, s => ushort.Parse(s, NumberStyles.Number, CultureInfo.CurrentCulture));
        AssertSameException<int>(text, s => int.Parse(s, NumberStyles.Number, CultureInfo.CurrentCulture));
        AssertSameException<uint>(text, s => uint.Parse(s, NumberStyles.Number, CultureInfo.CurrentCulture));
        AssertSameException<long>(text, s => long.Parse(s, NumberStyles.Number, CultureInfo.CurrentCulture));
        AssertSameException<ulong>(text, s => ulong.Parse(s, NumberStyles.Number, CultureInfo.CurrentCulture));
        AssertSameException<float>(text, s => float.Parse(s, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.CurrentCulture));
        AssertSameException<double>(text, s => double.Parse(s, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.CurrentCulture));
        AssertSameException<decimal>(text, s => decimal.Parse(s, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.CurrentCulture));
#if LIBRARY_NET8
        AssertSameException<Int128>(text, s => Int128.Parse(s, NumberStyles.Number, CultureInfo.CurrentCulture));
        AssertSameException<UInt128>(text, s => UInt128.Parse(s, NumberStyles.Number, CultureInfo.CurrentCulture));
        AssertSameException<Half>(text, s => Half.Parse(s, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.CurrentCulture));
#endif
    }

    [Theory]
    [InlineData("x")]
    [InlineData("1,5")]
    [InlineData("300")]
    [InlineData("1e400")]
    public void ReadRows_InvalidNumberWithStyle_ThrowsParseExceptionType(string text)
    {
        AssertSameException<int, IntegerStyleRow>(text, s => int.Parse(s, NumberStyles.Integer, CultureInfo.CurrentCulture));
        AssertSameException<byte, HexStyleRow>(text, s => byte.Parse(s, NumberStyles.HexNumber, CultureInfo.CurrentCulture));
        AssertSameException<decimal, CurrencyStyleRow>(text, s => decimal.Parse(s, NumberStyles.Currency, CultureInfo.CurrentCulture));
    }

    [Theory]
    [InlineData("x")]
    [InlineData("2024-13-45")]
    [InlineData("2024-01-02T10:00:00+15:00")]
    [InlineData("25:00")]
    public void ReadRows_InvalidDateOrTime_ThrowsParseExceptionType(string text)
    {
        AssertSameException<DateTime>(text, s => DateTime.ParseExact(s, ["G"], CultureInfo.CurrentCulture, DateTimeStyles.None));
        AssertSameException<DateTimeOffset>(text, s => DateTimeOffset.ParseExact(s, ["G"], CultureInfo.CurrentCulture, DateTimeStyles.None));
#if LIBRARY_NET8
        AssertSameException<DateOnly>(text, s => DateOnly.ParseExact(s, ["d"], CultureInfo.CurrentCulture, DateTimeStyles.None));
        AssertSameException<TimeOnly>(text, s => TimeOnly.ParseExact(s, ["t"], CultureInfo.CurrentCulture, DateTimeStyles.None));
#endif
    }

    [Theory]
    [InlineData("x")]
    [InlineData("{not-a-guid}")]
    public void ReadRows_InvalidGuidOrBool_ThrowsFormatException(string text)
    {
        AssertSameException<Guid>(text, s => Guid.Parse(s));
        AssertSameException<bool>(text, s => bool.Parse(s));
    }

    [Theory]
    [InlineData("Someday")]
    [InlineData("99999999999")]
    [InlineData("-99999999999")]
    [InlineData("Monday, x")]
    public void ReadRows_InvalidEnum_ThrowsParseExceptionType(string text)
    {
        AssertSameException<DayOfWeek>(text, s => Enum.Parse(typeof(DayOfWeek), s));
    }

    [Theory]
    [InlineData("http://[bad")]
    [InlineData("relative/path")]
    public void ReadRows_InvalidUri_ThrowsParseExceptionType(string text)
    {
        AssertSameException<Uri>(text, s => new Uri(s, UriKind.Absolute));
    }

    [Fact]
    public void ReadRows_InvalidValue_MessageNamesValueAndType()
    {
        using var importer = XlsxBuilder.Create(["Value"], ["abc"]);
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Row<int>>().ToArray());
        Assert.Contains("\"abc\"", exception.InnerException!.Message);
        Assert.Contains("Int32", exception.InnerException.Message);
    }

    [Fact]
    public void ReadRows_InvalidDate_MessageNamesFormats()
    {
        using var importer = XlsxBuilder.Create(["Value"], ["abc"]);
        importer.Configuration.RegisterClassMap<Row<DateTime>>(c => c.Map(r => r.Value).WithFormats("yyyy-MM-dd", "dd/MM/yyyy"));
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Row<DateTime>>().ToArray());
        Assert.Contains("\"yyyy-MM-dd\", \"dd/MM/yyyy\"", exception.InnerException!.Message);
    }

    private static void AssertSameException<T>(string text, Func<string, object> parse) => AssertSameException<T, Row<T>>(text, parse);

    private static void AssertSameException<T, TRow>(string text, Func<string, object> parse)
    {
        Exception? expected;
        try
        {
            parse(text);
            expected = null;
        }
        catch (Exception exception)
        {
            expected = exception;
        }

        using var importer = XlsxBuilder.Create(["Value"], [text]);
        var sheet = importer.ReadSheet();
        var results = sheet.ReadRowResults<TRow>().ToArray();
        var result = Assert.Single(results);
        if (expected is null)
        {
            // The text is valid for Parse; ExcelMapper may accept or reject it (for example 1.0 as an integer).
            return;
        }

        if (result.IsValid)
        {
            // ExcelMapper accepts some values Parse doesn't, such as formats tried by default.
            return;
        }

        var actual = Assert.Single(result.Errors).InnerException;
        Assert.True(expected.GetType() == actual?.GetType(), $"{typeof(T).Name} \"{text}\": expected {expected.GetType().Name}, got {actual?.GetType().Name ?? "null"}.");
    }

    public class Row<T>
    {
        public T Value { get; set; } = default!;
    }

    public class IntegerStyleRow
    {
        [ExcelNumberStyle(NumberStyles.Integer)]
        public int Value { get; set; }
    }

    public class HexStyleRow
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        public byte Value { get; set; }
    }

    public class CurrencyStyleRow
    {
        [ExcelNumberStyle(NumberStyles.Currency)]
        public decimal Value { get; set; }
    }
}
