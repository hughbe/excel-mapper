using System.Globalization;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// The built-in conversion of numbers, strings, dates, bools and GUIDs, with fallbacks, mapping
/// dictionaries and changes to a map after reading.
/// </summary>
public class MapBuiltInConversionTests
{
    [Fact]
    public void ReadRows_InvalidValue_CustomFallbackReceivesMappersException()
    {
        using var importer = Helpers.GetCsvImporter("Value\nabc\n");
        var fallback = new RecordingFallback();
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithInvalidFallback(fallback));
        Assert.Equal(-1, importer.ReadSheet().ReadRows<IntRow>().Single().Value);
        Assert.IsType<FormatException>(fallback.Exception);
    }

    private sealed class RecordingFallback : ICellFallback<int>
    {
        public Exception? Exception { get; private set; }

        public int Fallback(CellValue value, CellContext context, Exception? exception)
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
        // int.Parse reports a fractional part as an overflow.
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
    public void ReadRows_MappersChangedAfterReading_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\ntwo\n");
        var map = new ExcelClassMap<IntRow>();
        var valueMap = map.Map(r => r.Value);
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();
        Assert.Equal(1, sheet.ReadRow<IntRow>().Value);

        // The map is frozen once rows are read with it.
        Assert.Throws<InvalidOperationException>(() => valueMap.WithMapping(new Dictionary<string, int> { ["two"] = 2 }));
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntRow>());
    }

    [Fact]
    public void ReadRows_BuiltInThenMappingDictionary_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Value\nn/a\n3\nx\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value)
            .WithMapping(new Dictionary<string, int> { ["n/a"] = -2 })
            .WithInvalidFallback(0));
        Assert.Equal([-2, 3, 0], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
    }

    [Fact]
    public void ReadRows_EmptyValues_UseEmptyFallbacks()
    {
        // A blank row at the end of a file isn't a row, so a row with values follows it.
        using var importer = Helpers.GetCsvImporter("Value,Name,When\n,,\n1,a,2024-03-04\n");
        importer.Configuration.RegisterClassMap<Mixed>(c =>
        {
            c.Map(r => r.Value).WithEmptyFallback(5);
            c.Map(r => r.Name);
            c.Map(r => r.When).WithEmptyFallback(new DateTime(2024, 1, 1));
        });
        var row = importer.ReadSheet().ReadRows<Mixed>().First();
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
        using var importer = Helpers.GetCsvImporter($"Number,Flag,When,Id,Day\n5,yes,2024-03-04,{id},Monday\n,,,,\n5,yes,2024-03-04,{id},Monday\n");
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
        importer.Configuration.RegisterClassMap<EnumRow>(c => c.Map(r => r.Day).WithIgnoreCase());
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
        using var importer = Helpers.GetCsvImporter("Duration,Offset\n01:00:00,2024-03-04T10:30:00Z\n,\n01:00:00,2024-03-04T10:30:00Z\n");
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
        using var importer = Helpers.GetCsvImporter("Date,Time\n2024-03-04,14:30\n,\n2024-03-04,14:30\n");
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
    [InlineData(2.5, typeof(OverflowException))]
    public void ReadRows_NumberCellInvalidForType_Throws(double value, Type exceptionType)
    {
        using var importer = XlsxBuilder.Create(["Byte"], [value]);
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<ByteRow>().ToArray());
        Assert.Equal("A2", exception.CellAddress);
        Assert.IsType(exceptionType, exception.InnerException);
    }

    [Theory]
    [InlineData(typeof(sbyte), -128.0, 127.0, -129.0, 128.0)]
    [InlineData(typeof(byte), 0.0, 255.0, -1.0, 256.0)]
    [InlineData(typeof(short), -32768.0, 32767.0, -32769.0, 32768.0)]
    [InlineData(typeof(ushort), 0.0, 65535.0, -1.0, 65536.0)]
    [InlineData(typeof(int), -2147483648.0, 2147483647.0, -2147483649.0, 2147483648.0)]
    [InlineData(typeof(uint), 0.0, 4294967295.0, -1.0, 4294967296.0)]
    // The largest doubles below 2^63, 2^64 and 2^96, and those powers of two.
    [InlineData(typeof(long), -9223372036854775808.0, 9223372036854774784.0, -9223372036854777856.0, 9223372036854775808.0)]
    [InlineData(typeof(ulong), 0.0, 18446744073709549568.0, -1.0, 18446744073709551616.0)]
    [InlineData(typeof(decimal), -79228162514264328797450928128.0, 79228162514264328797450928128.0, -79228162514264337593543950336.0, 79228162514264337593543950336.0)]
    public void ReadRows_NumberCellAtTypeBounds_ReturnsExpected(Type type, double min, double max, double belowMin, double aboveMax)
    {
        using var importer = XlsxBuilder.Create(["Value"], [min], [max], [belowMin], [aboveMax]);
        var results = (IEnumerable<object>)typeof(MapBuiltInConversionTests).GetMethod(nameof(ReadResults), System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!
            .MakeGenericMethod(type).Invoke(null, [importer])!;
        Assert.Equal([true, true, false, false], results.Select(r => (bool)r.GetType().GetProperty("IsValid")!.GetValue(r)!));

        var errors = results.Skip(2).Select(r => ((IReadOnlyList<ExcelMappingException>)r.GetType().GetProperty("Errors")!.GetValue(r)!).Single());
        Assert.All(errors, e => Assert.IsType<OverflowException>(e.InnerException));
    }

    private static IEnumerable<object> ReadResults<T>(ExcelImporter importer)
        => importer.ReadSheet().ReadRowResults<ValueRow<T>>().Cast<object>().ToArray();

    public class ValueRow<T>
    {
        public T? Value { get; set; }
    }

    [Theory]
    [InlineData("", "2.5")]
    [InlineData("de-DE", "2,5")]
    public void ReadRows_FractionalNumberCellForInteger_MessageUsesImportersCulture(string culture, string number)
    {
        using var importer = XlsxBuilder.Create(["Byte"], [2.5]);
        importer.Configuration.Culture = new System.Globalization.CultureInfo(culture);
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<ByteRow>().ToArray());
        Assert.Equal($"The value {number} isn't a whole number, so it can't be byte.", exception.InnerException!.Message);
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
        using var importer = Helpers.GetCsvImporter("Int64,Decimal\n5,1.5\n,\n5,1.5\n");
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
    public void ReadRows_SplitElementsWithMappingDictionary_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Values,Dates\n1;two,\n");
        importer.Configuration.RegisterClassMap<ElementsRow>(c =>
        {
            c.Map(r => r.Values).WithSeparators(';').WithElementMap(e => e.WithMapping(new Dictionary<string, int> { ["two"] = 2 }));
            c.Map(r => r.Dates).WithSeparators(';');
        });
        Assert.Equal([1, 2], importer.ReadSheet().ReadRows<ElementsRow>().Single().Values);
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
