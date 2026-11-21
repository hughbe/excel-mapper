using System.Linq;
using ExcelMapper.Mappers;
using static ExcelMapper.Tests.XlsxBuilder;

namespace ExcelMapper.Tests;

public class DateTimeOffsetOffsetTests
{
    private static readonly DateTime s_date = new(2024, 3, 4, 5, 6, 7);

    public class Row
    {
        public DateTimeOffset When { get; set; }
    }

    public class NullableRow
    {
        public DateTimeOffset? When { get; set; }
    }

    public class UtcRow
    {
        [ExcelOffset("+00:00")]
        public DateTimeOffset When { get; set; }
    }

    public class NegativeRow
    {
        [ExcelOffset("-08:00")]
        public DateTimeOffset When { get; set; }
    }

    [Fact]
    public void Default_DateCell_UsesLocalOffset()
    {
        using var importer = Create(["When"], [s_date]);
        var when = importer.ReadSheet().ReadRows<Row>().Single().When;
        Assert.Equal(s_date, when.DateTime);
        Assert.Equal(TimeZoneInfo.Local.GetUtcOffset(s_date), when.Offset);
    }

    [Fact]
    public void WithOffset_DateCell_UsesOffset()
    {
        using var importer = Create(["When"], [s_date]);
        importer.Configuration.RegisterClassMap<Row>(c => c.Map(p => p.When).WithOffset(TimeSpan.FromHours(5.5)));

        var when = importer.ReadSheet().ReadRows<Row>().Single().When;
        Assert.Equal(new DateTimeOffset(s_date, TimeSpan.FromHours(5.5)), when);
        Assert.Equal(TimeSpan.FromHours(5.5), when.Offset);
    }

    [Fact]
    public void WithOffset_NullableAndSerialNumber_UsesOffset()
    {
        using var importer = Create(["When"], [s_date.ToOADate()], [null]);
        importer.Configuration.RegisterClassMap<NullableRow>(c => c.Map(p => p.When).WithOffset(TimeSpan.Zero));

        var rows = importer.ReadSheet().ReadRows<NullableRow>().ToArray();
        Assert.Equal(new DateTimeOffset(s_date, TimeSpan.Zero), rows[0].When);
        Assert.Null(rows[1].When);
    }

    [Fact]
    public void ExcelOffsetAttribute_UsesOffset()
    {
        using var utc = Create(["When"], [s_date]);
        Assert.Equal(new DateTimeOffset(s_date, TimeSpan.Zero), utc.ReadSheet().ReadRows<UtcRow>().Single().When);

        using var negative = Create(["When"], [s_date]);
        Assert.Equal(new DateTimeOffset(s_date, TimeSpan.FromHours(-8)), negative.ReadSheet().ReadRows<NegativeRow>().Single().When);
    }

    [Fact]
    public void WithOffset_TextWithOffset_KeepsTextOffset()
    {
        using var importer = Helpers.GetCsvImporter("When\n2024-03-04T05:06:07+02:00\n");
        importer.Configuration.RegisterClassMap<Row>(c => c.Map(p => p.When).WithOffset(TimeSpan.Zero).WithFormats("yyyy-MM-ddTHH:mm:sszzz"));

        Assert.Equal(TimeSpan.FromHours(2), importer.ReadSheet().ReadRows<Row>().Single().When.Offset);
    }

    [Theory]
    [InlineData("+00:00", 0)]
    [InlineData("00:00", 0)]
    [InlineData("+05:30", 330)]
    [InlineData("-08:00", -480)]
    [InlineData(" +14:00 ", 840)]
    public void ExcelOffsetAttribute_Ctor_ParsesOffset(string offset, int expectedMinutes)
    {
        Assert.Equal(TimeSpan.FromMinutes(expectedMinutes), new ExcelOffsetAttribute(offset).Offset);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Z")]
    [InlineData("5")]
    [InlineData("+15:00")]
    [InlineData("+05:60")]
    [InlineData("+5:30")]
    public void ExcelOffsetAttribute_InvalidOffset_ThrowsArgumentException(string offset)
    {
        Assert.Throws<ArgumentException>("offset", () => new ExcelOffsetAttribute(offset));
    }

    [Fact]
    public void ExcelOffsetAttribute_Null_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("offset", () => new ExcelOffsetAttribute(null!));
    }

    [Theory]
    [InlineData(15 * 60)]
    [InlineData(-15 * 60)]
    public void Offset_OutOfRange_ThrowsArgumentOutOfRangeException(int minutes)
    {
        var mapper = new DateTimeOffsetMapper();
        Assert.Throws<ArgumentOutOfRangeException>("value", () => mapper.Offset = TimeSpan.FromMinutes(minutes));
        Assert.Throws<ArgumentOutOfRangeException>("value", () => mapper.Offset = TimeSpan.FromSeconds(30));
        Assert.Null(mapper.Offset);
    }

    [Fact]
    public void Offset_SetAndReset_Success()
    {
        var mapper = new DateTimeOffsetMapper { Offset = TimeSpan.FromHours(-14) };
        Assert.Equal(TimeSpan.FromHours(-14), mapper.Offset);
        mapper.Offset = null;
        Assert.Null(mapper.Offset);
    }
}
