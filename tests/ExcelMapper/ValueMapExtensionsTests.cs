namespace ExcelMapper.Tests;

public class ValueMapExtensionsTests : ExcelClassMap<ValueMapExtensionsTests.Row>
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
    public void NullableMethods_NullArguments_ThrowArgumentNullException()
    {
        var map = Map(r => r.NullableNumber);
        Assert.Throws<ArgumentNullException>("converter", () => map.WithConverter((Func<string?, int?>)null!));
        Assert.Throws<ArgumentNullException>("converter", () => map.WithConverter((ICellConverter<int>)null!));
        Assert.Throws<ArgumentNullException>("fallback", () => map.WithEmptyFallback((ICellFallback<int>)null!));
        Assert.Throws<ArgumentNullException>("fallback", () => map.WithInvalidFallback((ICellFallback<int>)null!));
    }
}
