using System.Globalization;
using System.Linq;
using ExcelMapper.Abstractions;
using ExcelMapper.Mappers;

namespace ExcelMapper.Tests;

public class OneToOneMapExtensionsTests : ExcelClassMap<OneToOneMapExtensionsTests.Row>
{
    public class Row
    {
        public DateTime Date { get; set; }
        public DateTime? NullableDate { get; set; }
        public DateTimeOffset Offset { get; set; }
        public DateTimeOffset? NullableOffset { get; set; }
        public TimeSpan Span { get; set; }
        public TimeSpan? NullableSpan { get; set; }
#if LIBRARY_NET8
        public DateOnly Day { get; set; }
        public DateOnly? NullableDay { get; set; }
        public TimeOnly Time { get; set; }
        public TimeOnly? NullableTime { get; set; }
#endif
        public int Number { get; set; }
        public int? NullableNumber { get; set; }
    }

    [Fact]
    public void WithFormats_Chained_ReturnsMap()
    {
        // The methods return the map, so map methods such as MakeOptional chain after them.
        var map = Map(r => r.Date);
        Assert.Same(map, map.WithFormats("yyyy-MM-dd").WithFormatProvider(CultureInfo.InvariantCulture).MakeOptional());
        Assert.True(map.Optional);
        var mapper = map.Mappers.OfType<DateTimeMapper>().Single();
        Assert.Equal(["yyyy-MM-dd"], mapper.Formats);
        Assert.Same(CultureInfo.InvariantCulture, mapper.Provider);
    }

    [Fact]
    public void WithFormats_EveryDateAndTimeType_ReturnsMap()
    {
        AssertFormats(Map(r => r.Date), m => m.WithFormats("a"), m => m.WithFormats((IEnumerable<string>)["a"]));
        AssertFormats(Map(r => r.NullableDate), m => m.WithFormats("a"), m => m.WithFormats((IEnumerable<string>)["a"]));
        AssertFormats(Map(r => r.Offset), m => m.WithFormats("a"), m => m.WithFormats((IEnumerable<string>)["a"]));
        AssertFormats(Map(r => r.NullableOffset), m => m.WithFormats("a"), m => m.WithFormats((IEnumerable<string>)["a"]));
        AssertFormats(Map(r => r.Span), m => m.WithFormats("a"), m => m.WithFormats((IEnumerable<string>)["a"]));
        AssertFormats(Map(r => r.NullableSpan), m => m.WithFormats("a"), m => m.WithFormats((IEnumerable<string>)["a"]));
#if LIBRARY_NET8
        AssertFormats(Map(r => r.Day), m => m.WithFormats("a"), m => m.WithFormats((IEnumerable<string>)["a"]));
        AssertFormats(Map(r => r.NullableDay), m => m.WithFormats("a"), m => m.WithFormats((IEnumerable<string>)["a"]));
        AssertFormats(Map(r => r.Time), m => m.WithFormats("a"), m => m.WithFormats((IEnumerable<string>)["a"]));
        AssertFormats(Map(r => r.NullableTime), m => m.WithFormats("a"), m => m.WithFormats((IEnumerable<string>)["a"]));
#endif
    }

    private static void AssertFormats<T>(OneToOneMap<T> map, Func<OneToOneMap<T>, OneToOneMap<T>> withFormats, Func<OneToOneMap<T>, OneToOneMap<T>> withEnumerableFormats)
    {
        Assert.Same(map, withFormats(map));
        Assert.Equal(["a"], (string[])((dynamic)map.Mappers.Single(m => m.GetType().GetProperty("Formats") is not null)).Formats);
        Assert.Same(map, withEnumerableFormats(map));
        Assert.Same(map, map.MakeOptional());
    }

    [Fact]
    public void WithFormatProvider_EveryDateAndTimeType_ReturnsMap()
    {
        var provider = CultureInfo.InvariantCulture;
        Assert.Same(provider, Provider(Map(r => r.Date).WithFormatProvider(provider)));
        Assert.Same(provider, Provider(Map(r => r.NullableDate).WithFormatProvider(provider)));
        Assert.Same(provider, Provider(Map(r => r.Offset).WithFormatProvider(provider)));
        Assert.Same(provider, Provider(Map(r => r.NullableOffset).WithFormatProvider(provider)));
        Assert.Same(provider, Provider(Map(r => r.Span).WithFormatProvider(provider)));
        Assert.Same(provider, Provider(Map(r => r.NullableSpan).WithFormatProvider(provider)));
#if LIBRARY_NET8
        Assert.Same(provider, Provider(Map(r => r.Day).WithFormatProvider(provider)));
        Assert.Same(provider, Provider(Map(r => r.NullableDay).WithFormatProvider(provider)));
        Assert.Same(provider, Provider(Map(r => r.Time).WithFormatProvider(provider)));
        Assert.Same(provider, Provider(Map(r => r.NullableTime).WithFormatProvider(provider)));
#endif
    }

    private static IFormatProvider? Provider<T>(OneToOneMap<T> map)
        => ((dynamic)map.Mappers.Single(m => m.GetType().GetProperty("Provider") is not null)).Provider;

    [Fact]
    public void WithOffset_Chained_ReturnsMap()
    {
        var offset = TimeSpan.FromHours(2);
        var map = Map(r => r.Offset);
        Assert.Same(map, map.WithOffset(offset).MakeOptional());
        Assert.Equal(offset, map.Mappers.OfType<DateTimeOffsetMapper>().Single().Offset);

        var nullableMap = Map(r => r.NullableOffset);
        Assert.Same(nullableMap, nullableMap.WithOffset(offset).MakeOptional());
        Assert.Equal(offset, nullableMap.Mappers.OfType<DateTimeOffsetMapper>().Single().Offset);
    }

#if LIBRARY_NET8
    [Fact]
    public void WithNumberStyle_Chained_ReturnsMap()
    {
        var map = Map(r => r.Number);
        Assert.Same(map, map.WithNumberStyle(NumberStyles.HexNumber).WithFormatProvider(CultureInfo.InvariantCulture).MakeOptional());
        var mapper = map.Mappers.OfType<INumberBaseMapper<int>>().Single();
        Assert.Equal(NumberStyles.HexNumber, mapper.Style);
        Assert.Same(CultureInfo.InvariantCulture, mapper.Provider);

        var nullableMap = Map(r => r.NullableNumber);
        Assert.Same(nullableMap, nullableMap.WithNumberStyle(NumberStyles.HexNumber).WithFormatProvider(CultureInfo.InvariantCulture).MakeOptional());
        Assert.Equal(NumberStyles.HexNumber, nullableMap.Mappers.OfType<INumberBaseMapper<int>>().Single().Style);
    }
#endif

    [Fact]
    public void WithConverter_NullableWithUnderlyingTypeDelegate_Converts()
    {
        var map = Map(r => r.NullableNumber);
        Assert.Same(map, map.WithConverter(s => int.Parse(s!) * 2).MakeOptional());
        Assert.Equal(6, Assert.Single(map.Mappers).Map(new ReadCellResult(0, "3", preserveFormatting: false)).Value);
    }

    [Fact]
    public void WithConverter_NullableWithUnderlyingTypeConverter_Converts()
    {
        var map = Map(r => r.NullableNumber);
        Assert.Same(map, map.WithConverter(new FixedConverter(5)).MakeOptional());
        Assert.Equal(5, Assert.Single(map.Mappers).Map(new ReadCellResult(0, "3", preserveFormatting: false)).Value);
    }

    [Fact]
    public void WithFallbacks_NullableWithUnderlyingTypeFallback_SetsFallbacks()
    {
        var map = Map(r => r.NullableNumber);
        Assert.Same(map, map.WithEmptyFallback(new FixedFallback(1)).WithInvalidFallback(new FixedFallback(2)).MakeOptional());
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var sheet = importer.ReadSheet();
        Assert.Equal(1, map.EmptyFallback!.PerformFallback(sheet, 0, new ReadCellResult(0, "", preserveFormatting: false), null, null));
        Assert.Equal(2, map.InvalidFallback!.PerformFallback(sheet, 0, new ReadCellResult(0, "x", preserveFormatting: false), null, null));
    }

    [Fact]
    public void NullableMethods_NullArguments_ThrowArgumentNullException()
    {
        var map = Map(r => r.NullableNumber);
        Assert.Throws<ArgumentNullException>("converter", () => map.WithConverter((ConvertUsingSimpleMapperDelegate<int>)null!));
        Assert.Throws<ArgumentNullException>("converter", () => map.WithConverter((ICellConverter<int>)null!));
        Assert.Throws<ArgumentNullException>("fallback", () => map.WithEmptyFallback((ICellFallback<int>)null!));
        Assert.Throws<ArgumentNullException>("fallback", () => map.WithInvalidFallback((ICellFallback<int>)null!));
    }

    private sealed class FixedConverter(int value) : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue cell, CellContext context) => CellConversionResult<int>.Success(value);
    }

    private sealed class FixedFallback(int value) : ICellFallback<int>
    {
        public int Fallback(CellValue cell, CellContext context, Exception? exception) => value;
    }
}
