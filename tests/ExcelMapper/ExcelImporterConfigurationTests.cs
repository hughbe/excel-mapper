using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Tests;

public class ExcelImporterConfigurationTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SkipBlankLines_Set_GetReturnsExpected(bool value)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.SkipBlankLines = value;
        Assert.Equal(value, importer.Configuration.SkipBlankLines);

        // Set same.
        importer.Configuration.SkipBlankLines = value;
        Assert.Equal(value, importer.Configuration.SkipBlankLines);

        // Set different.
        importer.Configuration.SkipBlankLines = !value;
        Assert.Equal(!value, importer.Configuration.SkipBlankLines);
    }

    [Fact]
    public void MaxColumnsPerSheet_DefaultValue_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Equal(10000, importer.Configuration.MaxColumnsPerSheet);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(100)]
    [InlineData(1000)]
    [InlineData(16384)] // Excel maximum
    [InlineData(int.MaxValue)]
    public void MaxColumnsPerSheet_Set_GetReturnsExpected(int value)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.MaxColumnsPerSheet = value;
        Assert.Equal(value, importer.Configuration.MaxColumnsPerSheet);

        // Set same.
        importer.Configuration.MaxColumnsPerSheet = value;
        Assert.Equal(value, importer.Configuration.MaxColumnsPerSheet);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(0)]
    public void MaxColumnsPerSheet_SetInvalid_ThrowsArgumentOutOfRangeException(int value)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Throws<ArgumentOutOfRangeException>("value", () => importer.Configuration.MaxColumnsPerSheet = value);
    }

    [Fact]
    public void RegisterClassMap_InvokeDefault_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestMap>();

        Assert.True(importer.Configuration.TryGetClassMap<int>(out var classMap));
        var map = Assert.IsType<TestMap>(classMap);
        Assert.Equal(FallbackStrategy.ThrowIfPrimitive, map.EmptyValueStrategy);
        Assert.Equal(typeof(int), map.Type);
        Assert.Empty(map.Properties);

        Assert.True(importer.Configuration.TryGetClassMap(typeof(int), out classMap));
        map = Assert.IsType<TestMap>(classMap);
        Assert.Equal(FallbackStrategy.ThrowIfPrimitive, map.EmptyValueStrategy);
        Assert.Equal(typeof(int), map.Type);
        Assert.Empty(map.Properties);
    }

    [Fact]
    public void RegisterClassMap_InvokeAction_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<IntClass>(c =>
        {
            c.Map(o => o.Value1);
            c.Map(o => o.Value2);
        });

        Assert.True(importer.Configuration.TryGetClassMap<IntClass>(out var classMap));
        var map = Assert.IsType<ExcelClassMap<IntClass>>(classMap);
        Assert.Equal(FallbackStrategy.ThrowIfPrimitive, map.EmptyValueStrategy);
        Assert.Equal(typeof(IntClass), map.Type);
        Assert.Equal(2, map.Properties.Count);

        Assert.True(importer.Configuration.TryGetClassMap(typeof(IntClass), out classMap));
        map = Assert.IsType<ExcelClassMap<IntClass>>(classMap);
        Assert.Equal(FallbackStrategy.ThrowIfPrimitive, map.EmptyValueStrategy);
        Assert.Equal(typeof(IntClass), map.Type);
        Assert.Equal(2, map.Properties.Count);
    }

    [Fact]
    public void RegisterClassMap_InvokeActionEmpty_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<IntClass>(c =>
        {
        });

        Assert.True(importer.Configuration.TryGetClassMap<IntClass>(out var classMap));
        var map = Assert.IsType<ExcelClassMap<IntClass>>(classMap);
        Assert.Equal(FallbackStrategy.ThrowIfPrimitive, map.EmptyValueStrategy);
        Assert.Equal(typeof(IntClass), map.Type);
        Assert.Empty(map.Properties);

        Assert.True(importer.Configuration.TryGetClassMap(typeof(IntClass), out classMap));
        map = Assert.IsType<ExcelClassMap<IntClass>>(classMap);
        Assert.Equal(FallbackStrategy.ThrowIfPrimitive, map.EmptyValueStrategy);
        Assert.Equal(typeof(IntClass), map.Type);
        Assert.Empty(map.Properties);
    }

    private class IntClass
    {
        public int Value1 { get; set; }
        public int Value2 { get; set; }
    }

    [Fact]
    public void RegisterClassMap_NullClassMapFactory_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Throws<ArgumentNullException>("classMapFactory", () => importer.Configuration.RegisterClassMap<int>(null!));
    }

    [Fact]
    public void RegisterClassMap_InvokeExcelClassMap_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var map = new TestMap();
        importer.Configuration.RegisterClassMap(map);

        Assert.True(importer.Configuration.TryGetClassMap<int>(out var classMap));
        Assert.Same(map, classMap);

        Assert.True(importer.Configuration.TryGetClassMap(typeof(int), out classMap));
        Assert.Same(map, classMap);
    }

    [Fact]
    public void RegisterClassMap_InvokeTypeIMap_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var map = new CustomIMap();
        importer.Configuration.RegisterClassMap(typeof(int), map);

        Assert.True(importer.Configuration.TryGetClassMap<int>(out var classMap));
        Assert.Same(map, classMap);

        Assert.True(importer.Configuration.TryGetClassMap(typeof(int), out classMap));
        Assert.Same(map, classMap);
    }

    [Fact]
    public void RegisterClassMap_NullClassType_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var map = new CustomIMap();
        Assert.Throws<ArgumentNullException>("classType", () => importer.Configuration.RegisterClassMap(null!, map));
    }

    [Fact]
    public void RegisterClassMap_NullClassMap_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Throws<ArgumentNullException>("classMap", () => importer.Configuration.RegisterClassMap(null!));
        Assert.Throws<ArgumentNullException>("classMap", () => importer.Configuration.RegisterClassMap(typeof(int), null!));
    }

    [Fact]
    public void RegisterClassMap_ClassTypeAlreadyRegistered_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestMap>();
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<TestMap>());
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new TestMap()));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(int), new TestMap()));
    }

    [Fact]
    public void RegisterClassMap_ContainsPropertyWithMappers_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var map = new ValidMapperClassMap();
        importer.Configuration.RegisterClassMap(map);
        
        Assert.True(importer.Configuration.TryGetClassMap<RecordClass>(out var classMap));
        Assert.Same(map, classMap);
    }

    public record Id(int Value);

    public class RecordClass
    {
        public Id? Id { get; private set; }
    }

    public class ValidMapperClassMap : ExcelClassMap<RecordClass>
    {
        public ValidMapperClassMap()
        {
            Map(data => data.Id)
                .WithConverter(v => new Id(int.Parse(v!)))
                .WithColumnName("Value");
        }
    }

    [Fact]
    public void RegisterClassMap_ContainsPropertyWithoutMappers_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<NoMapperClassMap>());
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new NoMapperClassMap()));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(RecordClass), new NoMapperClassMap()));
    }

    public class NoMapperClassMap : ExcelClassMap<RecordClass>
    {
        public NoMapperClassMap()
        {
            Map(data => data.Id)
                .WithColumnName("Value");
        }
    }

    [Fact]
    public void RegisterClassMap_ArrayIndexerValueCantBeMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<IDisposableArrayClass>();
        map.Map(p => p.Value[0]);
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(ObjectArrayClass), map));
    }

    private class IDisposableArrayClass
    {
        public IDisposable[] Value { get; set; } = default!;
    }

    [Fact]
    public void RegisterClassMap_ContainsArrayIndexerPropertyWithoutMappers_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<DefaultObjectArrayIndexClassMap>());
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new DefaultObjectArrayIndexClassMap()));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(ObjectArrayClass), new DefaultObjectArrayIndexClassMap()));
    }

    private class DefaultObjectArrayIndexClassMap : ExcelClassMap<ObjectArrayClass>
    {
        public DefaultObjectArrayIndexClassMap()
        {
            Map(o => o.Values[0]);
            Map(o => o.Values[1]);
        }
    }

    private class ObjectArrayClass
    {
        public SimpleClass[] Values { get; set; } = default!;
    }

    private class SimpleClass
    {
        public int Value { get; set; }
    }

    [Fact]
    public void RegisterClassMap_ContainsMultidimensionalArrayValue_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClass>();
        map.Map(p => p.Value);
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(ObjectMultidimensionalClass), map));
    }

    private class MultidimensionalClass
    {
        public int[,] Value { get; set; } = default!;
    }

    [Fact]
    public void RegisterClassMap_ContainsMultidimensionalIndexerElementCantBeMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<IDisposableMultidimensionalClass>();
        map.Map(p => p.Value[0, 0]);
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(ObjectMultidimensionalClass), map));
    }

    private class IDisposableMultidimensionalClass
    {
        public IDisposable[,] Value { get; set; } = default!;
    }

    [Fact]
    public void RegisterClassMap_ContainsMultidimensionalIndexerPropertyWithoutMappers_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<DefaultObjectMultidimensionalIndexClassMap>());
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new DefaultObjectMultidimensionalIndexClassMap()));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(ObjectMultidimensionalClass), new DefaultObjectMultidimensionalIndexClassMap()));
    }

    private class DefaultObjectMultidimensionalIndexClassMap : ExcelClassMap<ObjectMultidimensionalClass>
    {
        public DefaultObjectMultidimensionalIndexClassMap()
        {
            Map(o => o.Values[0, 0]);
            Map(o => o.Values[1, 0]);
        }
    }

    private class ObjectMultidimensionalClass
    {
        public SimpleClass[,] Values { get; set; } = default!;
    }

    [Fact]
    public void RegisterClassMap_ContainsListIndexerElementCantBeMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<IDisposableListClass>();
        map.Map(p => p.Value[0]);
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(ObjectListClass), map));
    }

    private class IDisposableListClass
    {
        public List<IDisposable> Value { get; set; } = default!;
    }

    [Fact]
    public void RegisterClassMap_ContainsListIndexerPropertyWithoutMappers_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<DefaultObjectListIndexClassMap>());
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new DefaultObjectListIndexClassMap()));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(ObjectListClass), new DefaultObjectListIndexClassMap()));
    }

    private class DefaultObjectListIndexClassMap : ExcelClassMap<ObjectListClass>
    {
        public DefaultObjectListIndexClassMap()
        {
            Map(o => o.Values[0]);
            Map(o => o.Values[1]);
        }
    }

    private class ObjectListClass
    {
        public List<SimpleClass> Values { get; set; } = default!;
    }

    [Fact]
    public void RegisterClassMap_DictionaryIndexerValueCantBeMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<IDisposableDictionaryClass>();
        map.Map(p => p.Value["key"]);
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(ObjectDictionaryClass), map));
    }

    private class IDisposableDictionaryClass
    {
        public Dictionary<string, IDisposable> Value { get; set; } = default!;
    }

    [Fact]
    public void RegisterClassMap_ContainsDictionaryIndexerPropertyWithoutMappers_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<DefaultObjectDictionaryIndexClassMap>());
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new DefaultObjectDictionaryIndexClassMap()));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(ObjectDictionaryClass), new DefaultObjectDictionaryIndexClassMap()));
    }

    private class DefaultObjectDictionaryIndexClassMap : ExcelClassMap<ObjectDictionaryClass>
    {
        public DefaultObjectDictionaryIndexClassMap()
        {
            Map(o => o.Values["Key1"]);
            Map(o => o.Values["Key2"]);
        }
    }

    private class ObjectDictionaryClass
    {
        public Dictionary<string, SimpleClass> Values { get; set; } = default!;
    }

    [Fact]
    public void TryGetClassMap_NullClassType_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        IMap? classMap = null;
        Assert.Throws<ArgumentNullException>("classType", () => importer.Configuration.TryGetClassMap(null!, out classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryGetClassMap_NoSuchClassType_ReturnsFalse()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<OtherTestMap>();

        Assert.False(importer.Configuration.TryGetClassMap<TestMap>(out var classMap));
        Assert.Null(classMap);

        Assert.False(importer.Configuration.TryGetClassMap(typeof(TestMap), out classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryGetClassMap_NoRegisteredClassMaps_ReturnsFalse()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.False(importer.Configuration.TryGetClassMap<TestMap>(out var classMap));
        Assert.Null(classMap);

        Assert.False(importer.Configuration.TryGetClassMap(typeof(TestMap), out classMap));
        Assert.Null(classMap);
    }

    private class CustomIMap : IMap
    {
        public Type Type => typeof(int);

        public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
            => throw new NotImplementedException();
    }

    private class TestMap : ExcelClassMap<int>
    {
    }

    private class OtherTestMap : ExcelClassMap<int>
    {
    }
}
