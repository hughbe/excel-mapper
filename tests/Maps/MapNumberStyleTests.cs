using System.Globalization;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// Number parsing and <see cref="ExcelNumberStyleAttribute"/>, which behave the same on every build
/// of the library (run with -p:LibraryTargetFramework=netstandard2.0 to test the netstandard builds).
/// </summary>
public class MapNumberStyleTests
{
    public class Row<T>
    {
        public T Value { get; set; } = default!;
    }

    public class HexIntRow
    {
        [ExcelNumberStyle(NumberStyles.AllowHexSpecifier)]
        public int Value { get; set; }
    }

    public class HexNullableLongRow
    {
        [ExcelNumberStyle(NumberStyles.AllowHexSpecifier)]
        public long? Value { get; set; }
    }

    public class CurrencyRow
    {
        [ExcelNumberStyle(NumberStyles.Currency)]
        public decimal Value { get; set; }
    }

    public class HexDoubleRow
    {
        [ExcelNumberStyle(NumberStyles.AllowHexSpecifier)]
        public double Value { get; set; }
    }

    private static T Read<T>(string cell, CultureInfo? culture = null)
    {
        using var importer = Helpers.GetCsvImporter($"Value\n\"{cell}\"\n");
        importer.Configuration.Culture = culture ?? CultureInfo.InvariantCulture;
        return importer.ReadSheet().ReadRows<T>().Single();
    }

    private static ExcelMappingException ReadFails<T>(string cell)
    {
        using var importer = Helpers.GetCsvImporter($"Value\n\"{cell}\"\n");
        importer.Configuration.Culture = CultureInfo.InvariantCulture;
        var sheet = importer.ReadSheet();
        return Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<T>().ToArray());
    }

    [Theory]
    [InlineData("FF", 255)]
    [InlineData("0xFF", 255)]
    [InlineData("0Xff", 255)]
    [InlineData("&HFF", 255)]
    [InlineData("7FFFFFFF", int.MaxValue)]
    public void ReadRows_HexNumberStyleAttribute_ReturnsExpected(string cell, int expected)
    {
        Assert.Equal(expected, Read<HexIntRow>(cell).Value);
    }

    [Fact]
    public void ReadRows_HexNumberStyleAttributeNullable_ReturnsExpected()
    {
        Assert.Equal(0xABCL, Read<HexNullableLongRow>("0xABC").Value);
    }

    [Theory]
    [InlineData("12")]
    [InlineData("GG")]
    public void ReadRows_HexNumberStyleAttributeInvalid_ThrowsExcelMappingException(string cell)
    {
        if (cell == "12")
        {
            // "12" is valid hex.
            Assert.Equal(0x12, Read<HexIntRow>(cell).Value);
            return;
        }

        ReadFails<HexIntRow>(cell);
    }

    [Fact]
    public void ReadRows_HexNumberStyleAttributeFloatingPoint_ThrowsExcelMappingException()
    {
        // Hex isn't supported for floating-point types.
        var ex = ReadFails<HexDoubleRow>("FF");
        Assert.IsType<ArgumentException>(ex.InnerException);
    }

    [Fact]
    public void ReadRows_CurrencyNumberStyleAttribute_ReturnsExpected()
    {
        Assert.Equal(1234.50m, Read<CurrencyRow>("$1,234.50", new CultureInfo("en-US")).Value);
        Assert.Equal(-12m, Read<CurrencyRow>("($12.00)", new CultureInfo("en-US")).Value);
    }

    [Theory]
    [InlineData("1,234", 1234)]
    [InlineData(" 42 ", 42)]
    [InlineData("-5", -5)]
    [InlineData("5-", -5)]
    [InlineData("2.0", 2)]
    public void ReadRows_DefaultStyleInteger_ReturnsExpected(string cell, int expected)
    {
        Assert.Equal(expected, Read<Row<int>>(cell).Value);
    }

    [Theory]
    [InlineData("2.5")]
    [InlineData("1E5")]
    [InlineData("abc")]
    public void ReadRows_DefaultStyleIntegerInvalid_ThrowsExcelMappingException(string cell)
    {
        ReadFails<Row<int>>(cell);
    }

    [Theory]
    [InlineData("1E5", 100000.0)]
    [InlineData("1.5e-3", 0.0015)]
    [InlineData("1,234.5", 1234.5)]
    [InlineData("-2.5", -2.5)]
    public void ReadRows_DefaultStyleDouble_ReturnsExpected(string cell, double expected)
    {
        Assert.Equal(expected, Read<Row<double>>(cell).Value);
        Assert.Equal((float)expected, Read<Row<float>>(cell).Value);
    }

    [Fact]
    public void ReadRows_DefaultStyleDecimal_ReturnsExpected()
    {
        Assert.Equal(150m, Read<Row<decimal>>("1.5E2").Value);
        Assert.Equal(1234.5m, Read<Row<decimal>>("1,234.5").Value);
    }

    [Fact]
    public void ReadRows_OutOfRange_ThrowsExcelMappingException()
    {
        var ex = ReadFails<Row<byte>>("300");
        Assert.IsType<OverflowException>(ex.InnerException);
        ReadFails<Row<sbyte>>("-129");
        ReadFails<Row<uint>>("-1");
    }

    [Fact]
    public void ReadRows_Culture_ParsesWithCulture()
    {
        Assert.Equal(1234.5, Read<Row<double>>("1.234,5", new CultureInfo("de-DE")).Value);
    }

    [Fact]
    public void ReadRows_AllIntegerTypes_ReturnsExpected()
    {
        Assert.Equal((sbyte)-1, Read<Row<sbyte>>("-1").Value);
        Assert.Equal((byte)255, Read<Row<byte>>("255").Value);
        Assert.Equal((short)-300, Read<Row<short>>("-300").Value);
        Assert.Equal((ushort)65535, Read<Row<ushort>>("65,535").Value);
        Assert.Equal(-70000, Read<Row<int>>("-70000").Value);
        Assert.Equal(4000000000u, Read<Row<uint>>("4,000,000,000").Value);
        Assert.Equal(-5000000000L, Read<Row<long>>("-5000000000").Value);
        Assert.Equal(18000000000000000000UL, Read<Row<ulong>>("18000000000000000000").Value);
    }

    [Fact]
    public void ReadRows_Nullable_ReturnsExpected()
    {
        Assert.Equal(5, Read<Row<int?>>("5").Value);
        Assert.Equal(1.5, Read<Row<double?>>("1.5").Value);
    }
}
