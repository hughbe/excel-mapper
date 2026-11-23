using System.Linq;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Tests;

public class ExcelImporterConfigurationTests
{
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
    public void SkipBlankLines_GetDefault_ReturnsFalse()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.False(importer.Configuration.ValidateDataAnnotations);
    }

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
    public void ValidateDataAnnotations_GetDefault_ReturnsFalse()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.False(importer.Configuration.ValidateDataAnnotations);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ValidateDataAnnotations_Set_GetReturnsExpected(bool value)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.ValidateDataAnnotations = value;
        Assert.Equal(value, importer.Configuration.ValidateDataAnnotations);

        // Set same.
        importer.Configuration.ValidateDataAnnotations = value;
        Assert.Equal(value, importer.Configuration.ValidateDataAnnotations);

        // Set different.
        importer.Configuration.ValidateDataAnnotations = !value;
        Assert.Equal(!value, importer.Configuration.ValidateDataAnnotations);
    }

    [Fact]
    public void RegisterClassMap_InvokeDefault_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestMap>();

        Assert.True(importer.Configuration.TryGetClassMap<int>(out var classMap));
        var map = Assert.IsType<TestMap>(classMap);
        Assert.Equal(EmptyValueStrategy.Error, map.EmptyValueStrategy);
        Assert.Equal(typeof(int), map.Type);

        Assert.True(importer.Configuration.TryGetClassMap(typeof(int), out var untypedClassMap));
        map = Assert.IsType<TestMap>(untypedClassMap);
        Assert.Equal(EmptyValueStrategy.Error, map.EmptyValueStrategy);
        Assert.Equal(typeof(int), map.Type);
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
        Assert.Equal(EmptyValueStrategy.Error, map.EmptyValueStrategy);
        Assert.Equal(typeof(IntClass), map.Type);

        Assert.True(importer.Configuration.TryGetClassMap(typeof(IntClass), out var untypedClassMap));
        map = Assert.IsType<ExcelClassMap<IntClass>>(untypedClassMap);
        Assert.Equal(EmptyValueStrategy.Error, map.EmptyValueStrategy);
        Assert.Equal(typeof(IntClass), map.Type);
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
        Assert.Equal(EmptyValueStrategy.Error, map.EmptyValueStrategy);
        Assert.Equal(typeof(IntClass), map.Type);

        Assert.True(importer.Configuration.TryGetClassMap(typeof(IntClass), out var untypedClassMap));
        map = Assert.IsType<ExcelClassMap<IntClass>>(untypedClassMap);
        Assert.Equal(EmptyValueStrategy.Error, map.EmptyValueStrategy);
        Assert.Equal(typeof(IntClass), map.Type);
    }

    private class IntClass
    {
        public int Value1 { get; set; }
        public int Value2 { get; set; }
    }

    [Fact]
    public void RegisterClassMap_NullConfigure_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Throws<ArgumentNullException>("configure", () => importer.Configuration.RegisterClassMap<int>(null!));
    }

    [Fact]
    public void RegisterClassMap_InvokeExcelClassMap_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var map = new TestMap();
        importer.Configuration.RegisterClassMap(map);

        Assert.True(importer.Configuration.TryGetClassMap<int>(out var classMap));
        Assert.Same(map, classMap);

        Assert.True(importer.Configuration.TryGetClassMap(typeof(int), out var untypedClassMap));
        Assert.Same(map, untypedClassMap);
    }

    [Fact]
    public void RegisterClassMap_NullClassMap_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.Throws<ArgumentNullException>("classMap", () => importer.Configuration.RegisterClassMap(null!));
    }

    // Derives from the non-generic ExcelClassMap, so it can't map any members.
    [Fact]
    public void RegisterClassMap_DerivedFromExcelClassMapOfT_Success()
    {
        // A class map that derives from a class map of T, not ExcelClassMap<T> directly.
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var map = new DerivedTestMap();
        importer.Configuration.RegisterClassMap(map);
        Assert.True(importer.Configuration.TryGetClassMap<int>(out var classMap));
        Assert.Same(map, classMap);
    }

    private class DerivedTestMap : TestMap
    {
    }

    [Fact]
    public void RegisterClassMap_ClassTypeAlreadyRegistered_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<TestMap>();
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<TestMap>());
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new TestMap()));
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
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new NoMapperClassMap()));
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
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
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
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new DefaultObjectArrayIndexClassMap()));
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
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
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
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
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
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new DefaultObjectMultidimensionalIndexClassMap()));
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
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
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
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new DefaultObjectListIndexClassMap()));
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
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
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
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new DefaultObjectDictionaryIndexClassMap()));
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
        ExcelClassMap? classMap = null;
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

        Assert.False(importer.Configuration.TryGetClassMap(typeof(TestMap), out var untypedClassMap));
        Assert.Null(untypedClassMap);
    }

    [Fact]
    public void TryGetClassMap_AutoMapped_ReturnsAutoMappedClassMap()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();
        sheet.ReadRow<int>();
        sheet.ReadRow<IntValueClass>();

        Assert.True(importer.Configuration.TryGetClassMap<int>(out var classMap));
        Assert.Equal(typeof(int), classMap.Type);

        Assert.True(importer.Configuration.TryGetClassMap<IntValueClass>(out var valueClassMap));
        Assert.Equal(typeof(IntValueClass), valueClassMap.Type);
    }

    private class IntValueClass
    {
        [ExcelColumnName("Int Value")]
        public int Value { get; set; }
    }

    [Fact]
    public void TryGetClassMap_NoRegisteredClassMaps_ReturnsFalse()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        Assert.False(importer.Configuration.TryGetClassMap<TestMap>(out var classMap));
        Assert.Null(classMap);

        Assert.False(importer.Configuration.TryGetClassMap(typeof(TestMap), out var untypedClassMap));
        Assert.Null(untypedClassMap);
    }

    private class TestMap : ExcelClassMap<int>
    {
    }

    private class OtherTestMap : ExcelClassMap<int>
    {
    }

    [Fact]
    public void IgnoreSeparatorsInColumnNames_GetDefault_ReturnsFalse()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        Assert.False(importer.Configuration.IgnoreSeparatorsInColumnNames);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void IgnoreSeparatorsInColumnNames_Set_GetReturnsExpected(bool value)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.IgnoreSeparatorsInColumnNames = value;
        Assert.Equal(value, importer.Configuration.IgnoreSeparatorsInColumnNames);

        // Set same.
        importer.Configuration.IgnoreSeparatorsInColumnNames = value;
        Assert.Equal(value, importer.Configuration.IgnoreSeparatorsInColumnNames);

        // Set different.
        importer.Configuration.IgnoreSeparatorsInColumnNames = !value;
        Assert.Equal(!value, importer.Configuration.IgnoreSeparatorsInColumnNames);
    }

    [Fact]
    public void Culture_GetDefault_ReturnsInvariantCulture()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        Assert.Same(System.Globalization.CultureInfo.InvariantCulture, importer.Configuration.Culture);
    }

    public static IEnumerable<object?[]> Culture_Set_TestData()
    {
        yield return new object?[] { System.Globalization.CultureInfo.InvariantCulture };
        yield return new object?[] { new System.Globalization.CultureInfo("de-DE") };
    }

    [Theory]
    [MemberData(nameof(Culture_Set_TestData))]
    public void Culture_Set_GetReturnsExpected(System.Globalization.CultureInfo value)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.Culture = value;
        Assert.Same(value, importer.Configuration.Culture);

        // Set same.
        importer.Configuration.Culture = value;
        Assert.Same(value, importer.Configuration.Culture);
    }

    [Fact]
    public void Culture_SetNull_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        Assert.Throws<ArgumentNullException>("value", () => importer.Configuration.Culture = null!);
    }

    public class AutoMappedClass
    {
        public string Value { get; set; } = default!;
    }

    [Fact]
    public void RegisterClassMap_AfterAutoMapping_ReplacesAutoMappedClassMap()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadRow<AutoMappedClass>();
        Assert.True(importer.Configuration.TryGetClassMap<AutoMappedClass>(out var autoMappedClassMap));

        var classMap = new ExcelClassMap<AutoMappedClass>();
        classMap.Map(p => p.Value);
        importer.Configuration.RegisterClassMap(classMap);
        Assert.True(importer.Configuration.TryGetClassMap<AutoMappedClass>(out var registeredClassMap));
        Assert.Same(classMap, registeredClassMap);
        Assert.NotSame(autoMappedClassMap, registeredClassMap);

        // A class map registered by the user can't be replaced.
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new ExcelClassMap<AutoMappedClass>()));
    }

    [Fact]
    public void RegisterClassMap_RowMapWithoutMappers_ThrowsExcelMappingException()
    {
        // A class map read as a single value is checked like its members.
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var map = new ExcelClassMap<RowWithoutMappers>();
        map.Map(p => p);
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
    }

    public class RowWithoutMappers
    {
        public int A { get; set; }
    }

    [Theory]
    [MemberData(nameof(RegisterClassMap_NoConversion_TestData))]
    public void RegisterClassMap_MemberWithNoConversion_ThrowsExcelMappingException(Action<ExcelImporterConfiguration> register, string message)
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var exception = Assert.Throws<ExcelMappingException>(() => register(importer.Configuration));
        Assert.Equal(message, exception.Message);
    }

    public static IEnumerable<object[]> RegisterClassMap_NoConversion_TestData()
    {
        yield return [(Action<ExcelImporterConfiguration>)(c => c.RegisterClassMap<NoConversionRow>(m => m.Map(r => r.Tags))), "Member \"Tags\" of type HashSet<string> is a collection that Map can't read from a cell: map it with MapCollection<string>(...)."];
        yield return [(Action<ExcelImporterConfiguration>)(c => c.RegisterClassMap<NoConversionRow>(m => m.Map(r => r.Scores))), "Member \"Scores\" of type SortedDictionary<string, int> is a dictionary that Map can't read from a cell: map it with MapDictionary<int>(...)."];
        yield return [(Action<ExcelImporterConfiguration>)(c => c.RegisterClassMap<NoConversionRow>(m => m.Map(r => r.Address))), "Member \"Address\" of type ExcelImporterConfigurationTests.NoConversionAddress has no conversion from a cell. Give it one with WithConverter or [ExcelConverter], or map its own members with MapObject(...)."];
    }

    [Fact]
    public void ReadRows_UnregisteredClassMapMemberWithNoConversion_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Tags\na\n");
        var classMap = new ExcelClassMap<NoConversionRow>();
        classMap.Map(r => r.Tags);
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows(classMap).ToArray());
        Assert.StartsWith("Member \"Tags\" of type HashSet<string> is a collection", exception.Message);
    }

    [Fact]
    public void RegisterClassMap_MapCollectionForHashSet_ReadsCollection()
    {
        using var importer = Helpers.GetCsvImporter("Tags\n\"a,b\"\n");
        importer.Configuration.RegisterClassMap<NoConversionRow>(m => m.MapCollection<string>(r => r.Tags));
        Assert.Equal(["a", "b"], importer.ReadSheet().ReadRows<NoConversionRow>().Single().Tags.Order());
    }

    public class NoConversionRow
    {
        public HashSet<string> Tags { get; set; } = [];

        public SortedDictionary<string, int> Scores { get; set; } = [];

        public NoConversionAddress Address { get; set; } = new();
    }

    public class NoConversionAddress
    {
        public string? Street { get; set; }
    }
}
