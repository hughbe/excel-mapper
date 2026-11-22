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

    // The tests run with GenericTypes.RequireRegisteredTypes on (see RequireRegisteredTypes.cs), so
    // a context throws if it maps a type it didn't register.
    private sealed class MissingContext<TRow>(Action<ExcelTypeRegistrar> register) : ExcelMappingContext
    {
        protected override void RegisterTypes(ExcelTypeRegistrar types)
        {
            types.Row<TRow>();
            register(types);
        }
    }

    public class NullableRow
    {
        public int? Value { get; set; }
    }

    public enum MissingEnum { A, B }

    public class EnumRow
    {
        public MissingEnum Value { get; set; }
    }

    public class ListRow
    {
        [ExcelSeparators(';')]
        public List<long> Values { get; set; } = [];
    }

    [Theory]
    [MemberData(nameof(MissingRegistration_TestData))]
    public void GetClassMap_TypeNotRegistered_ThrowsInvalidOperationException(Func<ExcelMappingContext> createContext, Func<ExcelMappingContext, object> getClassMap, string missing)
    {
        var context = createContext();
        var exception = Assert.Throws<InvalidOperationException>(() => getClassMap(context));
        Assert.Contains(missing, exception.Message);
        Assert.Contains(context.GetType().ToString(), exception.Message);
    }

    public static IEnumerable<object[]> MissingRegistration_TestData()
    {
        // A member's type.
        yield return [() => new MissingContext<Row>(t => t.Type<string>()), (ExcelMappingContext c) => c.GetClassMap<Row>(), "the type System.Int32"];

        // A nullable type, and its underlying type registered without it.
        yield return [() => new MissingContext<NullableRow>(t => t.Type<int>()), (ExcelMappingContext c) => c.GetClassMap<NullableRow>(), "System.Nullable`1[System.Int32]"];

        // An enum registered as a type: its mapper isn't registered.
        yield return [() => new MissingContext<EnumRow>(t => t.Type<MissingEnum>()), (ExcelMappingContext c) => c.GetClassMap<EnumRow>(), "the enum"];

        // A collection without its element type's collections.
        yield return [() => new MissingContext<ListRow>(t => { t.Type<List<long>>(); t.Type<long>(); }), (ExcelMappingContext c) => c.GetClassMap<ListRow>(), "collections of System.Int64"];

        // A dictionary row without its key and value types' dictionaries.
        yield return [() => new MissingContext<Dictionary<string, long>>(t => t.Type<long>()), (ExcelMappingContext c) => c.GetClassMap<Dictionary<string, long>>(), "dictionaries of System.String to System.Int64"];
    }

    [Fact]
    public void GetClassMap_EveryTypeRegistered_DoesNotThrow()
    {
        var context = new MissingContext<ListRow>(t => { t.Collection<List<long>, long>(); t.Type<long>(); });
        using var importer = Helpers.GetCsvImporter("Values\n\"1;2\"\n");
        Assert.Equal([1L, 2L], importer.ReadSheet().ReadRows(context.GetClassMap<ListRow>()).Single().Values);
    }

    [Fact]
    public void ReadRows_AutoMappedTypeNotRegistered_DoesNotThrow()
    {
        // Only a context's class maps are checked.
        using var importer = Helpers.GetCsvImporter("Value\nB\n");
        Assert.Equal(MissingEnum.B, importer.ReadSheet().ReadRows<EnumRow>().Single().Value);
    }
}
