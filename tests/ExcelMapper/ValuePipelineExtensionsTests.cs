using System.Globalization;
using System.Linq;
using System.Numerics;

namespace ExcelMapper.Tests;

public class ValuePipelineExtensionsTests : ExcelClassMap<Helpers.TestClass>
{
    public static IEnumerable<object[]> Formats_TestData()
    {
        yield return new object[] { new string[] { "1" } };
        yield return new object[] { new string[] { "g", "yyyy-MM-dd" } };
        yield return new object[] { new List<string> { "g", "yyyy-MM-dd" } };
    }

    [Fact]
    public void WithFormats_DateTimeNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.DateTimeValue);
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_DateTimeEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_DateTimeNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_DateTimeEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Fact]
    public void WithFormats_DateTimeNullableNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.NullableDateTimeValue);
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_DateTimeNullableEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_DateTimeNullableNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_DateTimeNullableEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.DateTimeOffsetValue);
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeOffsetValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeOffsetValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateTimeOffsetValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullableNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.NullableDateTimeOffsetValue);
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullableEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeOffsetValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullableNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeOffsetValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_DateTimeOffsetNullableEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateTimeOffsetValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Fact]
    public void WithFormats_TimeSpanNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.TimeSpanValue);
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_TimeSpanEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeSpanValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_TimeSpanNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeSpanValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_TimeSpanEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeSpanValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

    [Fact]
    public void WithFormats_TimeSpanNullableNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.NullableTimeSpanValue);
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithFormats_TimeSpanNullableEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeSpanValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }

    [Fact]
    public void WithFormats_TimeSpanNullableNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeSpanValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }

    [Fact]
    public void WithFormats_TimeSpanNullableEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeSpanValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_DateOnlyNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.DateOnlyValue);
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_DateOnlyEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_DateOnlyNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_DateOnlyEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.DateOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_DateOnlyNullableNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.NullableDateOnlyValue);
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_DateOnlyNullableEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_DateOnlyNullableNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_DateOnlyNullableEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableDateOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_TimeOnlyNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.TimeOnlyValue);
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_TimeOnlyEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_TimeOnlyNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_TimeOnlyEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.TimeOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_TimeOnlyNullableNullFormats_ThrowsArgumentNullException()
    {
        var map = Map(t => t.NullableTimeOnlyValue);
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats(null!));
        Assert.Throws<ArgumentNullException>("formats", () => map.WithFormats((IEnumerable<string>)null!));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_TimeOnlyNullableEmptyFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string>()));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_TimeOnlyNullableNullValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([null!]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { null! }));
    }
#endif

#if LIBRARY_NET8
    [Fact]
    public void WithFormats_TimeOnlyNullableEmptyValueInFormats_ThrowsArgumentException()
    {
        var map = Map(t => t.NullableTimeOnlyValue);
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats([string.Empty]));
        Assert.Throws<ArgumentException>("formats", () => map.WithFormats(new List<string> { string.Empty }));
    }
#endif

    public static IEnumerable<object?[]> FormatProvider_TestData()
    {
        yield return new object?[] { null };
        yield return new object[] { new CultureInfo("fr-FR") };
    }

    [Theory]
    [InlineData(UriKind.RelativeOrAbsolute - 1)]
    [InlineData(UriKind.Relative + 1)]
    public void WithUriKind_InvalidUriKind_ThrowsArgumentOutOfRangeException(UriKind uriKind)
    {
        var map = Map(t => t.UriValue);
        Assert.Throws<ArgumentOutOfRangeException>("uriKind", () => map.WithUriKind(uriKind));
    }

    [Fact]
    public void WithFallback_NullCellFallback_ThrowsArgumentNullException()
    {
        var map = Map(t => t.DateTimeValue);
        Assert.Throws<ArgumentNullException>("fallback", () => map.WithEmptyFallback((ICellFallback<DateTime>)null!));
        Assert.Throws<ArgumentNullException>("fallback", () => map.WithInvalidFallback((ICellFallback<DateTime>)null!));

        // A fallback for the underlying type of a nullable member.
        var nullableMap = Map(t => t.NullableDateTimeValue);
        Assert.Throws<ArgumentNullException>("fallback", () => nullableMap.WithEmptyFallback((ICellFallback<DateTime>)null!));
        Assert.Throws<ArgumentNullException>("fallback", () => nullableMap.WithInvalidFallback((ICellFallback<DateTime>)null!));
    }

    [Fact]
    public void Methods_NullMap_ThrowArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("map", () => ((ValueMap<int>)null!).WithConverter(new ZeroConverter()));
        Assert.Throws<ArgumentNullException>("map", () => ((ValueMap<int?>)null!).WithConverter(new ZeroConverter()));
        Assert.Throws<ArgumentNullException>("map", () => ((ValueMap<int>)null!).WithEmptyFallback(new ZeroFallback()));
        Assert.Throws<ArgumentNullException>("map", () => ((ValueMap<int?>)null!).WithEmptyFallback(new ZeroFallback()));
        Assert.Throws<ArgumentNullException>("map", () => ((ValueMap<int>)null!).WithInvalidFallback(new ZeroFallback()));
        Assert.Throws<ArgumentNullException>("map", () => ((ValueMap<int?>)null!).WithInvalidFallback(new ZeroFallback()));
    }

    [Fact]
    public void Methods_ReturnTheMap()
    {
        var map = Map(t => t.DateTimeValue);
        Assert.Same(map, map.WithFormats("yyyy-MM-dd"));
        Assert.Same(map, map.WithFormatProvider(CultureInfo.InvariantCulture));
        Assert.Same(map, map.WithEmptyFallback(new MinValueFallback()));
        Assert.Same(map, map.WithInvalidFallback(new MinValueFallback()));
        Assert.Same(map, map.WithConverter(new MinValueConverter()));

        var nullableMap = Map(t => t.NullableDateTimeValue);
        Assert.Same(nullableMap, nullableMap.WithFormats("yyyy-MM-dd"));
        Assert.Same(nullableMap, nullableMap.WithEmptyFallback(new MinValueFallback()));
        Assert.Same(nullableMap, nullableMap.WithInvalidFallback(new MinValueFallback()));
        Assert.Same(nullableMap, nullableMap.WithConverter(new MinValueConverter()));

        var uriMap = new ExcelClassMap<UriClass>().Map(u => u.Uri);
        Assert.Same(uriMap, uriMap.WithUriKind(UriKind.Relative));
    }

    private sealed class UriClass
    {
        public Uri Uri { get; set; } = default!;
    }

    private sealed class ZeroConverter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context) => CellConversionResult<int>.Success(0);
    }

    private sealed class ZeroFallback : ICellFallback<int>
    {
        public int Fallback(CellValue value, CellContext context, Exception? exception) => 0;
    }

    private sealed class MinValueConverter : ICellConverter<DateTime>
    {
        public CellConversionResult<DateTime> Convert(CellValue value, CellContext context) => CellConversionResult<DateTime>.Success(DateTime.MinValue);
    }

    private sealed class MinValueFallback : ICellFallback<DateTime>
    {
        public DateTime Fallback(CellValue value, CellContext context, Exception? exception) => DateTime.MinValue;
    }

    [Fact]
    public void WithConverter_NullCellConverter_ThrowsArgumentNullException()
    {
        var map = Map(t => t.DateTimeValue);
        Assert.Throws<ArgumentNullException>("converter", () => map.WithConverter((ICellConverter<DateTime>)null!));
        var nullableMap = Map(t => t.NullableTimeSpanValue);
        Assert.Throws<ArgumentNullException>("converter", () => nullableMap.WithConverter((ICellConverter<TimeSpan>)null!));
    }

    private sealed class FixedConverter<T>(T value) : ICellConverter<T>
    {
        public CellConversionResult<T> Convert(CellValue cell, CellContext context) => CellConversionResult<T>.Success(value);
    }

    [Fact]
    public void WithOffset_InvalidOffset_ThrowsArgumentOutOfRangeException()
    {
        var map = Map(t => t.DateTimeOffsetValue);
        Assert.Throws<ArgumentOutOfRangeException>("value", () => map.WithOffset(TimeSpan.FromHours(15)));
    }

    [Fact]
    public void WithOffset_NullMap_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("map", () => ((ValueMap<DateTimeOffset>)null!).WithOffset(TimeSpan.Zero));
        Assert.Throws<ArgumentNullException>("map", () => ((ValueMap<DateTimeOffset?>)null!).WithOffset(TimeSpan.Zero));
    }
}
