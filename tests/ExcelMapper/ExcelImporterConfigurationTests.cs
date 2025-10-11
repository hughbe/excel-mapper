using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using Xunit;

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
    public void RegisterClassMap_InvokeDefault_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestMap>();

        Assert.True(importer.Configuration.TryGetClassMap<int>(out var classMap));
        TestMap map = Assert.IsType<TestMap>(classMap);
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

    // When we bring non-string dictionary keys back, re-enable this test.
#if false
    [Fact]
    public void RegisterClassMap_NonStringDictionaryKey_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<IntDictionaryClass>();
        map.Map(p => p.Value[0]);
        using var importer = Helpers.GetImporter("Numbers.xlsx");
        importer.Configuration.RegisterClassMap(map);

        var sheet = importer.ReadSheet();
        var row = sheet.ReadRow<IntDictionaryClass>();
    }

    private class IntDictionaryClass
    {
        public Dictionary<int, string> Value { get; set; } = default!;
    }
#endif

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
