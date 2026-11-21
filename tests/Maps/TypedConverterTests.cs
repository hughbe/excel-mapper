using System.Globalization;
using System.Linq;
using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Mappers;

namespace ExcelMapper.Tests;

/// <summary>
/// Members whose mappers are all built-in typed converters (strings, numbers, dates, bools, GUIDs)
/// are mapped without boxing. These check that the typed path behaves exactly like the pipeline.
/// </summary>
public class TypedConverterTests
{
    [Fact]
    public void ReadRows_InvalidValue_CustomFallbackReceivesMappersException()
    {
        using var importer = Helpers.GetCsvImporter("Value\nabc\n");
        var fallback = new RecordingFallback();
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithInvalidFallbackItem(fallback));
        Assert.Equal(-1, importer.ReadSheet().ReadRows<IntRow>().Single().Value);
        Assert.IsType<FormatException>(fallback.Exception);
    }

    private sealed class RecordingFallback : IFallbackItem
    {
        public Exception? Exception { get; private set; }

        public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
        {
            Exception = exception;
            return -1;
        }
    }

    [Fact]
    public void ReadRows_InvalidValue_ErrorUnchanged()
    {
        using var importer = Helpers.GetCsvImporter("Value\n2.5\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<IntRow>().ToArray());
        Assert.Equal("A2", exception.CellAddress);
        Assert.Contains("Cannot assign \"2.5\" to member \"Value\"", exception.Message);
        // As before the typed path: int.Parse reports a fractional part as an overflow.
        Assert.IsType<OverflowException>(exception.InnerException);
    }

    [Fact]
    public void ReadRows_InvalidValueWithFixedFallback_ReturnsFallback()
    {
        using var importer = Helpers.GetCsvImporter("Value\nabc\n7\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithInvalidFallback(-1));
        Assert.Equal([-1, 7], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
    }

    [Fact]
    public void ReadRows_MappersChangedAfterReading_UsesNewMappers()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n2\n");
        var map = new ExcelClassMap<IntRow>();
        var valueMap = map.Map(r => r.Value);
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();
        Assert.Equal(1, sheet.ReadRow<IntRow>().Value);

        // A custom mapper isn't a typed converter: the member goes back to the pipeline.
        valueMap.WithMappers(new DoublingMapper());
        Assert.Equal(4, sheet.ReadRow<IntRow>().Value);
    }

    private sealed class DoublingMapper : ICellMapper
    {
        public CellMapperResult Map(ReadCellResult readResult) => CellMapperResult.Success(int.Parse(readResult.GetString()!) * 2);
    }

#if LIBRARY_NET8
    [Fact]
    public void ReadRows_CustomMapperThenBuiltIn_UsesPipeline()
    {
        // A custom mapper that ignores some values, then the built-in one.
        using var importer = Helpers.GetCsvImporter("Value\nn/a\n3\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value)
            .WithMappers(new IgnoreNotApplicable(), new INumberBaseMapper<int>())
            .WithInvalidFallback(0));
        Assert.Equal([0, 3], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
    }

    private sealed class IgnoreNotApplicable : ICellMapper
    {
        public CellMapperResult Map(ReadCellResult readResult)
            => readResult.GetString() == "n/a" ? CellMapperResult.Invalid(new FormatException()) : CellMapperResult.Ignore();
    }
#endif

    [Fact]
    public void ReadRows_EmptyValues_UseEmptyFallbacks()
    {
        using var importer = Helpers.GetCsvImporter("Value,Name,When\n,,\n");
        importer.Configuration.RegisterClassMap<Mixed>(c =>
        {
            c.Map(r => r.Value).WithEmptyFallback(5);
            c.Map(r => r.Name);
            c.Map(r => r.When).WithEmptyFallback(new DateTime(2024, 1, 1));
        });
        var row = importer.ReadSheet().ReadRows<Mixed>().Single();
        Assert.Equal(5, row.Value);
        Assert.Null(row.Name);
        Assert.Equal(new DateTime(2024, 1, 1), row.When);
    }

    [Theory]
    [InlineData("yes", true)]
    [InlineData("No", false)]
    [InlineData("1", true)]
    [InlineData("0", false)]
    [InlineData("TRUE", true)]
    public void ReadRows_Bool_ReturnsExpected(string text, bool expected)
    {
        using var importer = Helpers.GetCsvImporter($"Value\n{text}\n");
        Assert.Equal(expected, importer.ReadSheet().ReadRows<BoolRow>().Single().Value);
    }

    [Fact]
    public void ReadRows_TypedCellsXlsx_ReturnsExpected()
    {
        var date = new DateTime(2024, 3, 4);
        using var importer = XlsxBuilder.Create(["Value", "Name", "When"], [2.0, "a", date], [3.0, "b", 45000.0]);
        var rows = importer.ReadSheet().ReadRows<Mixed>().ToArray();
        Assert.Equal([2, 3], rows.Select(r => r.Value));
        Assert.Equal(date, rows[0].When);
        Assert.Equal(new DateTime(2023, 3, 15), rows[1].When);
    }

    [Fact]
    public void ReadRows_NullableMembers_ReturnsExpected()
    {
        var id = Guid.NewGuid();
        using var importer = Helpers.GetCsvImporter($"Number,Flag,When,Id,Day\n5,yes,2024-03-04,{id},Monday\n,,,,\n");
        var rows = importer.ReadSheet().ReadRows<NullableRow>().ToArray();
        Assert.Equal(5, rows[0].Number);
        Assert.True(rows[0].Flag);
        Assert.Equal(new DateTime(2024, 3, 4), rows[0].When);
        Assert.Equal(id, rows[0].Id);
        Assert.Equal(DayOfWeek.Monday, rows[0].Day);
        Assert.Null(rows[1].Number);
        Assert.Null(rows[1].Flag);
        Assert.Null(rows[1].When);
        Assert.Null(rows[1].Id);
        Assert.Null(rows[1].Day);
    }

    [Theory]
    [InlineData("Number", "x", "A2")]
    [InlineData("Flag", "maybe", "B2")]
    [InlineData("When", "nope", "C2")]
    [InlineData("Id", "x", "D2")]
    [InlineData("Day", "Someday", "E2")]
    public void ReadRows_NullableMemberInvalid_Throws(string column, string value, string address)
    {
        var values = new[] { "Number", "Flag", "When", "Id", "Day" }.Select(c => c == column ? value : "").ToArray();
        using var importer = Helpers.GetCsvImporter("Number,Flag,When,Id,Day\n" + string.Join(",", values) + "\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<NullableRow>().ToArray());
        Assert.Equal(address, exception.CellAddress);
    }

    public class NullableRow
    {
        public int? Number { get; set; }
        public bool? Flag { get; set; }
        public DateTime? When { get; set; }
        public Guid? Id { get; set; }
        public DayOfWeek? Day { get; set; }
    }

    [Theory]
    [InlineData("Monday", DayOfWeek.Monday)]
    [InlineData("1", DayOfWeek.Monday)]
    [InlineData("99", (DayOfWeek)99)]
    public void ReadRows_Enum_ReturnsExpected(string text, DayOfWeek expected)
    {
        using var importer = Helpers.GetCsvImporter($"Day\n{text}\n");
        Assert.Equal(expected, importer.ReadSheet().ReadRows<EnumRow>().Single().Day);
    }

    [Fact]
    public void ReadRows_EnumIgnoreCase_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Day\nmonday\n");
        importer.Configuration.RegisterClassMap<EnumRow>(c => c.Map(r => r.Day, ignoreCase: true));
        Assert.Equal(DayOfWeek.Monday, importer.ReadSheet().ReadRows<EnumRow>().Single().Day);
    }

    [Fact]
    public void ReadRows_EnumCaseSensitiveByDefault_Throws()
    {
        using var importer = Helpers.GetCsvImporter("Day\nmonday\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<EnumRow>().ToArray());
        Assert.IsType<ArgumentException>(exception.InnerException);
    }

    [Fact]
    public void ReadRows_DateAndTimeText_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Duration,Offset\n01:30:00,2024-03-04T10:30:00+01:00\n");
        var row = importer.ReadSheet().ReadRows<DateAndTimeRow>().Single();
        Assert.Equal(TimeSpan.FromMinutes(90), row.Duration);
        Assert.Equal(new DateTimeOffset(2024, 3, 4, 10, 30, 0, TimeSpan.FromHours(1)), row.Offset);
    }

    [Fact]
    public void ReadRows_DateAndTimeXlsx_ReturnsExpected()
    {
        // Typed cells, and plain numbers: a number of days and a serial date.
        using var importer = XlsxBuilder.Create(
            ["Duration", "Offset"],
            [new XlsxBuilder.Duration(TimeSpan.FromHours(30)), new DateTime(2024, 3, 4, 10, 30, 0)],
            [0.25, 45000.0]);
        importer.Configuration.RegisterClassMap<DateAndTimeRow>(c =>
        {
            c.Map(r => r.Duration);
            c.Map(r => r.Offset).WithOffset(TimeSpan.FromHours(2));
        });
        var rows = importer.ReadSheet().ReadRows<DateAndTimeRow>().ToArray();
        Assert.Equal(TimeSpan.FromHours(30), rows[0].Duration);
        Assert.Equal(new DateTimeOffset(2024, 3, 4, 10, 30, 0, TimeSpan.FromHours(2)), rows[0].Offset);
        Assert.Equal(TimeSpan.FromHours(6), rows[1].Duration);
        Assert.Equal(new DateTimeOffset(2023, 3, 15, 0, 0, 0, TimeSpan.FromHours(2)), rows[1].Offset);
    }

    [Fact]
    public void ReadRows_DateTimeOffsetOutOfRangeForOffset_Throws()
    {
        // The latest Excel date with a negative offset is after DateTimeOffset.MaxValue in UTC.
        using var importer = XlsxBuilder.Create(["Offset"], [new DateTime(9999, 12, 31, 23, 0, 0)]);
        importer.Configuration.RegisterClassMap<DateAndTimeRow>(c => c.Map(r => r.Offset).WithOffset(TimeSpan.FromHours(-14)));
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<DateAndTimeRow>().ToArray());
        Assert.Equal("A2", exception.CellAddress);
        Assert.IsType<ArgumentOutOfRangeException>(exception.InnerException);
    }

    [Fact]
    public void ReadRows_NullableDateAndTime_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Duration,Offset\n01:00:00,2024-03-04T10:30:00Z\n,\n");
        var rows = importer.ReadSheet().ReadRows<NullableDateAndTimeRow>().ToArray();
        Assert.Equal(TimeSpan.FromHours(1), rows[0].Duration);
        Assert.Equal(new DateTimeOffset(2024, 3, 4, 10, 30, 0, TimeSpan.Zero), rows[0].Offset);
        Assert.Null(rows[1].Duration);
        Assert.Null(rows[1].Offset);
    }

    [Theory]
    [InlineData("Duration", "A2")]
    [InlineData("Offset", "B2")]
    public void ReadRows_NullableDateAndTimeInvalid_Throws(string column, string address)
    {
        using var importer = Helpers.GetCsvImporter("Duration,Offset\n" + (column == "Duration" ? "x," : ",x") + "\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<NullableDateAndTimeRow>().ToArray());
        Assert.Equal(address, exception.CellAddress);
        Assert.IsType<FormatException>(exception.InnerException);
    }

    public class DateAndTimeRow
    {
        public TimeSpan Duration { get; set; }
        public DateTimeOffset Offset { get; set; }
    }

    public class NullableDateAndTimeRow
    {
        public TimeSpan? Duration { get; set; }
        public DateTimeOffset? Offset { get; set; }
    }

#if LIBRARY_NET8
    [Fact]
    public void ReadRows_DateOnlyAndTimeOnly_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Date,Time\n2024-03-04,14:30\n,\n");
        var rows = importer.ReadSheet().ReadRows<DateOnlyRow>().ToArray();
        Assert.Equal(new DateOnly(2024, 3, 4), rows[0].Date);
        Assert.Equal(new TimeOnly(14, 30), rows[0].Time);
        Assert.Null(rows[1].Date);
        Assert.Null(rows[1].Time);
    }

    [Fact]
    public void ReadRows_DateOnlyAndTimeOnlyXlsx_ReturnsExpected()
    {
        using var importer = XlsxBuilder.Create(["Date", "Time"], [new DateOnly(2024, 3, 4), new TimeOnly(9, 15)], [45000.0, 0.25]);
        var rows = importer.ReadSheet().ReadRows<DateOnlyRow>().ToArray();
        Assert.Equal(new DateOnly(2024, 3, 4), rows[0].Date);
        Assert.Equal(new TimeOnly(9, 15), rows[0].Time);
        Assert.Equal(new DateOnly(2023, 3, 15), rows[1].Date);
        Assert.Equal(new TimeOnly(6, 0), rows[1].Time);
    }

    [Theory]
    [InlineData("Date", "A2")]
    [InlineData("Time", "B2")]
    public void ReadRows_DateOnlyAndTimeOnlyInvalid_Throws(string column, string address)
    {
        using var importer = Helpers.GetCsvImporter("Date,Time\n" + (column == "Date" ? "x," : ",x") + "\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<DateOnlyRow>().ToArray());
        Assert.Equal(address, exception.CellAddress);
    }

    public class DateOnlyRow
    {
        public DateOnly? Date { get; set; }
        public TimeOnly? Time { get; set; }
    }
#endif

    [Fact]
    public void ReadRows_EveryNumberTypeText_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("SByte,Byte,Int16,UInt16,Int32,UInt32,Int64,UInt64,Single,Double,Decimal\n-1,2,-3,4,\"-5,000\",6,-7,8,9.5,1e3,11.25\n");
        var row = importer.ReadSheet().ReadRows<AllNumbers>().Single();
        Assert.Equal(new AllNumbers { SByte = -1, Byte = 2, Int16 = -3, UInt16 = 4, Int32 = -5000, UInt32 = 6, Int64 = -7, UInt64 = 8, Single = 9.5f, Double = 1000, Decimal = 11.25m }, row);
    }

    [Fact]
    public void ReadRows_EveryNumberTypeXlsx_ReturnsExpected()
    {
        using var importer = XlsxBuilder.Create(
            ["SByte", "Byte", "Int16", "UInt16", "Int32", "UInt32", "Int64", "UInt64", "Single", "Double", "Decimal"],
            [-1.0, 2.0, -3.0, 4.0, -5.0, 6.0, -7.0, 8.0, 9.5, 10.5, 11.25]);
        var row = importer.ReadSheet().ReadRows<AllNumbers>().Single();
        Assert.Equal(new AllNumbers { SByte = -1, Byte = 2, Int16 = -3, UInt16 = 4, Int32 = -5, UInt32 = 6, Int64 = -7, UInt64 = 8, Single = 9.5f, Double = 10.5, Decimal = 11.25m }, row);
    }

    [Theory]
    [InlineData(300.0, typeof(OverflowException))]
    [InlineData(-1.0, typeof(OverflowException))]
    [InlineData(2.5, typeof(FormatException))]
    public void ReadRows_NumberCellInvalidForType_Throws(double value, Type exceptionType)
    {
        using var importer = XlsxBuilder.Create(["Byte"], [value]);
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<ByteRow>().ToArray());
        Assert.Equal("A2", exception.CellAddress);
        Assert.IsType(exceptionType, exception.InnerException);
    }

    [Fact]
    public void ReadRows_NumberCellInvalidWithFallback_ReturnsFallback()
    {
        using var importer = XlsxBuilder.Create(["Byte"], [300.0], [7.0]);
        importer.Configuration.RegisterClassMap<ByteRow>(c => c.Map(r => r.Byte).WithInvalidFallback((byte)1));
        Assert.Equal([1, 7], importer.ReadSheet().ReadRows<ByteRow>().Select(r => (int)r.Byte));
    }

    [Fact]
    public void ReadRows_NumberStyle_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Int32,Byte\n0x1F,FF\n");
        var row = importer.ReadSheet().ReadRows<HexNumbers>().Single();
        Assert.Equal(31, row.Int32);
        Assert.Equal(255, row.Byte);
    }

    public class HexNumbers
    {
        [ExcelNumberStyle(NumberStyles.HexNumber)]
        public int Int32 { get; set; }

        [ExcelNumberStyle(NumberStyles.HexNumber)]
        public byte Byte { get; set; }
    }

    [Fact]
    public void ReadRows_NullableNumbers_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Int64,Decimal\n5,1.5\n,\n");
        var rows = importer.ReadSheet().ReadRows<NullableNumbers>().ToArray();
        Assert.Equal(5, rows[0].Int64);
        Assert.Equal(1.5m, rows[0].Decimal);
        Assert.Null(rows[1].Int64);
        Assert.Null(rows[1].Decimal);
    }

    [Fact]
    public void ReadRows_NullableNumberInvalid_Throws()
    {
        using var importer = Helpers.GetCsvImporter("Int64,Decimal\n,x\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<NullableNumbers>().ToArray());
        Assert.Equal("B2", exception.CellAddress);
        Assert.IsType<FormatException>(exception.InnerException);
    }

    public record AllNumbers
    {
        public sbyte SByte { get; set; }
        public byte Byte { get; set; }
        public short Int16 { get; set; }
        public ushort UInt16 { get; set; }
        public int Int32 { get; set; }
        public uint UInt32 { get; set; }
        public long Int64 { get; set; }
        public ulong UInt64 { get; set; }
        public float Single { get; set; }
        public double Double { get; set; }
        public decimal Decimal { get; set; }
    }

    public class ByteRow
    {
        public byte Byte { get; set; }
    }

    public class NullableNumbers
    {
        public long? Int64 { get; set; }
        public decimal? Decimal { get; set; }
    }

    [Fact]
    public void ReadRows_SplitElements_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Values,Dates\n1;2;3,2024-01-02;2024-03-04\n");
        var row = importer.ReadSheet().ReadRows<ElementsRow>().Single();
        Assert.Equal([1, 2, 3], row.Values);
        Assert.Equal([new DateTime(2024, 1, 2), new DateTime(2024, 3, 4)], row.Dates);
    }

    [Fact]
    public void ReadRows_SplitElementInvalid_ThrowsAtCell()
    {
        using var importer = Helpers.GetCsvImporter("Values,Dates\n1;x;3,\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<ElementsRow>().ToArray());
        Assert.Equal("A2", exception.CellAddress);
        Assert.IsType<FormatException>(exception.InnerException);
    }

    [Fact]
    public void ReadRows_SplitElementsWithFallbacks_ReturnsFallbacks()
    {
        using var importer = Helpers.GetCsvImporter("Values,Dates\n1;x;;4,\n");
        importer.Configuration.RegisterClassMap<ElementsRow>(c =>
        {
            c.Map(r => r.Values).WithSeparators(';').WithElementMap(e => e.WithInvalidFallback(-1).WithEmptyFallback(0));
            c.Map(r => r.Dates).WithSeparators(';');
        });
        Assert.Equal([1, -1, 0, 4], importer.ReadSheet().ReadRows<ElementsRow>().Single().Values);
    }

    [Fact]
    public void ReadRows_SplitElementsWithCustomMapper_UsesPipeline()
    {
        using var importer = Helpers.GetCsvImporter("Values,Dates\n1;2,\n");
        importer.Configuration.RegisterClassMap<ElementsRow>(c =>
        {
            c.Map(r => r.Values).WithSeparators(';').WithElementMap(e => e.WithMappers(new DoublingMapper()));
            c.Map(r => r.Dates).WithSeparators(';');
        });
        Assert.Equal([2, 4], importer.ReadSheet().ReadRows<ElementsRow>().Single().Values);
    }

    [Fact]
    public void ReadRows_ColumnsToListAndDictionary_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("A,B,C\n1,2.5,3\n");
        var row = importer.ReadSheet().ReadRows<ColumnsRow>().Single();
        Assert.Equal([1, 3], row.Values);
        Assert.Equal(new Dictionary<string, decimal> { ["A"] = 1, ["B"] = 2.5m }, row.Amounts);
    }

    [Fact]
    public void ReadRows_DictionaryValueInvalid_ThrowsAtCell()
    {
        using var importer = Helpers.GetCsvImporter("A,B,C\n1,x,3\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<ColumnsRow>().ToArray());
        Assert.Equal("B2", exception.CellAddress);
    }

    public class ElementsRow
    {
        [ExcelSeparators(';')]
        public int[] Values { get; set; } = [];

        [ExcelSeparators(';')]
        public List<DateTime> Dates { get; set; } = [];
    }

    public class ColumnsRow
    {
        [ExcelColumnNames("A", "C")]
        public List<int> Values { get; set; } = [];

        [ExcelColumnNames("A", "B")]
        public Dictionary<string, decimal> Amounts { get; set; } = [];
    }

    public class EnumRow
    {
        public DayOfWeek Day { get; set; }
    }

    public class IntRow
    {
        public int Value { get; set; }
    }

    public class BoolRow
    {
        public bool Value { get; set; }
    }

    public class Mixed
    {
        public int Value { get; set; }
        public string? Name { get; set; }
        public DateTime When { get; set; }
    }
}
