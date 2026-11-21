namespace ExcelMapper.Tests;

public class RowNumberMapTests
{
    public static IEnumerable<object[]> Ctor_Type_TestData()
    {
        yield return new object[] { typeof(int) };
        yield return new object[] { typeof(int?) };
        yield return new object[] { typeof(long) };
        yield return new object[] { typeof(long?) };
        yield return new object[] { typeof(sbyte) };
        yield return new object[] { typeof(byte) };
        yield return new object[] { typeof(short) };
        yield return new object[] { typeof(ushort) };
        yield return new object[] { typeof(uint) };
        yield return new object[] { typeof(ulong) };
        yield return new object[] { typeof(float) };
        yield return new object[] { typeof(double) };
        yield return new object[] { typeof(decimal) };
        yield return new object[] { typeof(decimal?) };
    }

    [Theory]
    [MemberData(nameof(Ctor_Type_TestData))]
    public void Ctor_Type(Type type)
    {
        var map = new RowNumberMap(type);
        Assert.Same(type, map.Type);
    }

    [Fact]
    public void Ctor_NullType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("type", () => new RowNumberMap(null!));
    }

    [Theory]
    [InlineData(typeof(string))]
    [InlineData(typeof(object))]
    [InlineData(typeof(bool))]
    [InlineData(typeof(char))]
    [InlineData(typeof(DateTime))]
    [InlineData(typeof(DayOfWeek))]
    [InlineData(typeof(DayOfWeek?))]
    [InlineData(typeof(Guid))]
    public void Ctor_InvalidType_ThrowsArgumentException(Type type)
    {
        Assert.Throws<ArgumentException>("type", () => new RowNumberMap(type));
    }

    public static IEnumerable<object[]> TryGetValue_TestData()
    {
        yield return new object[] { typeof(int), 0, 1 };
        yield return new object[] { typeof(int), 9, 10 };
        yield return new object[] { typeof(int?), 9, 10 };
        yield return new object[] { typeof(int), int.MaxValue - 1, int.MaxValue };
        yield return new object[] { typeof(long), 9, 10L };
        yield return new object[] { typeof(long), int.MaxValue, 2147483648L };
        yield return new object[] { typeof(long?), 9, 10L };
        yield return new object[] { typeof(sbyte), 126, (sbyte)127 };
        yield return new object[] { typeof(byte), 254, (byte)255 };
        yield return new object[] { typeof(short), 9, (short)10 };
        yield return new object[] { typeof(ushort), 9, (ushort)10 };
        yield return new object[] { typeof(uint), 9, 10u };
        yield return new object[] { typeof(ulong), 9, 10ul };
        yield return new object[] { typeof(float), 9, 10f };
        yield return new object[] { typeof(double), 9, 10d };
        yield return new object[] { typeof(decimal), 9, 10m };
        yield return new object[] { typeof(decimal?), 9, 10m };
    }

    [Theory]
    [MemberData(nameof(TryGetValue_TestData))]
    public void TryGetValue_Invoke_ReturnsRowNumber(Type type, int rowIndex, object expected)
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var sheet = importer.ReadSheet();
        var map = new RowNumberMap(type);
        Assert.True(map.TryGetValue(sheet, rowIndex, importer.Reader, null, out var value));
        Assert.Equal(expected, value);
        Assert.IsType(expected.GetType(), value);
    }

    [Theory]
    [InlineData(typeof(sbyte), 127)]
    [InlineData(typeof(byte), 255)]
    [InlineData(typeof(short), short.MaxValue)]
    [InlineData(typeof(ushort), ushort.MaxValue)]
    public void TryGetValue_RowNumberTooLarge_ThrowsExcelMappingException(Type type, int rowIndex)
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var sheet = importer.ReadSheet();
        var map = new RowNumberMap(type);
        var member = typeof(RowNumberMapTests).GetMethod(nameof(TryGetValue_RowNumberTooLarge_ThrowsExcelMappingException))!;
        var exception = Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, rowIndex, importer.Reader, member, out _));
        Assert.IsType<OverflowException>(exception.InnerException);
        Assert.Equal(rowIndex, exception.RowIndex);
        Assert.Same(member, exception.Member);
        Assert.Contains($"Row number {rowIndex + 1L} is too large", exception.Message);
    }

    [Fact]
    public void TryGetValue_NullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var map = new RowNumberMap(typeof(int));
        Assert.Throws<ArgumentNullException>("sheet", () => map.TryGetValue(null!, 0, importer.Reader, null, out _));
    }
}
