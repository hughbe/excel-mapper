using System.Linq;

namespace ExcelMapper.Tests;

public class ExcelMappingContextTests
{
    private sealed class Context : ExcelMappingContext
    {
        public int RegisterCount { get; private set; }

        protected override void RegisterTypes(ExcelTypeRegistrar types)
        {
            RegisterCount++;
            types.Row<Row>();
            types.Type<string>();
            types.Type<int>();
            types.Row<IDisposable>();
        }
    }

    public class Row
    {
        public string? Name { get; set; }
        public int Age { get; set; }
    }

    [Fact]
    public void GetClassMap_RegisteredRow_AutoMapsType()
    {
        var context = new Context();
        using var importer = Helpers.GetCsvImporter("Name,Age\na,1\nb,2\n");
        var rows = importer.ReadSheet().ReadRows(context.GetClassMap<Row>()).ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
        Assert.Equal([1, 2], rows.Select(r => r.Age));
    }

    [Fact]
    public void GetClassMap_CalledTwice_ReturnsSameClassMapAndRegistersOnce()
    {
        var context = new Context();
        var classMap = context.GetClassMap<Row>();
        Assert.Same(classMap, context.GetClassMap<Row>());
        Assert.Equal(1, context.RegisterCount);
    }

    [Fact]
    public void GetClassMap_Invoke_ReturnsFrozenClassMap()
    {
        var classMap = new Context().GetClassMap<Row>();
        Assert.Throws<InvalidOperationException>(() => classMap.Map(r => r.Name));
    }

    [Fact]
    public void GetClassMap_UnregisteredType_ThrowsInvalidOperationException()
    {
        var context = new Context();
        Assert.Throws<InvalidOperationException>(() => context.GetClassMap<string>());
    }

    [Fact]
    public void GetClassMap_TypeCantBeAutoMapped_ThrowsExcelMappingException()
    {
        var context = new Context();
        Assert.Throws<ExcelMappingException>(() => context.GetClassMap<IDisposable>());
    }

    [Fact]
    public void GetClassMap_ManyThreads_RegistersOnce()
    {
        var context = new Context();
        var classMaps = Enumerable.Range(0, 8).AsParallel().Select(_ => context.GetClassMap<Row>()).ToArray();
        Assert.All(classMaps, classMap => Assert.Same(classMaps[0], classMap));
        Assert.Equal(1, context.RegisterCount);
    }
}
