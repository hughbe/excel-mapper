namespace ExcelMapper.Factories;

public class DictionaryFactoryTests
{
    [Fact]
    public void Begin_End_Success()
    {
        var factory = new DictionaryFactory<string, int>();

        // Begin.
        factory.Begin(1);
        var value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal([], value);

        // Begin again.
        factory.Begin(1);
        value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Begin_AlreadyBegan_ThrowsExcelMappingException()
    {
        var factory = new DictionaryFactory<string, int>();
        factory.Begin(1);
        Assert.Throws<ExcelMappingException>(() => factory.Begin(1));
    }

    [Fact]
    public void Begin_NegativeCount_ThrowsArgumentOutOfRangeException()
    {
        var factory = new DictionaryFactory<string, int>();
        Assert.Throws<ArgumentOutOfRangeException>("count", () => factory.Begin(-1));
    }

    [Fact]
    public void Add_End_Success()
    {
        var factory = new DictionaryFactory<string, int>();

        // Begin.
        factory.Begin(1);
        factory.Add("key", 1);
        var value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal(new Dictionary<string, int> { ["key"] = 1 }, value);

        // Begin again.
        factory.Begin(1);
        factory.Add("key", 2);
        value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal(new Dictionary<string, int> { ["key"] = 2 }, value);
    }

    [Fact]
    public void Add_OutOfRange_Success()
    {
        var factory = new DictionaryFactory<string, int>();
        factory.Begin(1);
        factory.Add("key1", 2);

        factory.Add("key2", 3);
        
        var value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal(new Dictionary<string, int> { ["key1"] = 2, ["key2"] = 3 }, value);
    }

    [Fact]
    public void Add_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new DictionaryFactory<string, int>();
        Assert.Throws<ExcelMappingException>(() => factory.Add("key", 1));
    }

    [Fact]
    public void Set_Invoke_Success()
    {
        var factory = new DictionaryFactory<string, int>();
        factory.Begin(1);
        factory.Add("key1", 1);

        Assert.Equal(new Dictionary<string, int> { ["key1"] = 1 }, Assert.IsType<Dictionary<string, int>>(factory.End()));
    }

    [Fact]
    public void Set_InvokeOutOfRange_Success()
    {
        var factory = new DictionaryFactory<string, int>();
        factory.Begin(1);
        factory.Add("key1", 1);
        factory.Add("key2", 2);

        Assert.Equal(new Dictionary<string, int> { ["key1"] = 1, ["key2"] = 2 }, Assert.IsType<Dictionary<string, int>>(factory.End()));
    }

    [Fact]
    public void Set_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new DictionaryFactory<string, int>();
        Assert.Throws<ExcelMappingException>(() => factory.Add("key", 1));
    }

    [Fact]
    public void Add_NullKey_ThrowsArgumentNullException()
    {
        var factory = new DictionaryFactory<string, int>();
        factory.Begin(1);
        Assert.Throws<ArgumentNullException>("key", () => factory.Add(null!, 1));
    }

    [Fact]
    public void Add_MultipleTimes_ThrowsArgumentException()
    {
        var factory = new DictionaryFactory<string, int>();
        factory.Begin(1);
        factory.Add("key", 1);

        Assert.Throws<ArgumentException>(null, () => factory.Add("key", 2));
    }

    [Fact]
    public void End_NotBegan_ThrowsExcelMappingException()
    {
        var factory = new DictionaryFactory<string, int>();
        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void End_AlreadyEnded_ThrowsExcelMappingException()
    {
        var factory = new DictionaryFactory<string, int>();
        factory.Begin(1);
        factory.End();

        Assert.Throws<ExcelMappingException>(() => factory.End());
    }

    [Fact]
    public void Reset_Invoke_Success()
    {
        var factory = new DictionaryFactory<string, int>();
        factory.Begin(1);
        factory.End();

        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }

    [Fact]
    public void Reset_NotBegan_Success()
    {
        var factory = new DictionaryFactory<string, int>();
        factory.Reset();

        // Make sure we can begin.
        factory.Begin(1);
        var value = Assert.IsType<Dictionary<string, int>>(factory.End());
        Assert.Equal([], value);
    }
}
