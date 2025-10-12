using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Dynamic;
using System.Linq;
using Xunit;

namespace ExcelMapper.Utilities.Tests;

public class AutoMapperTests
{
    [Theory]
    [InlineData(FallbackStrategy.ThrowIfPrimitive)]
    [InlineData(FallbackStrategy.SetToDefaultValue)]
    public void TryCreateClass_Map_ValidType_ReturnsTrue(FallbackStrategy emptyValueStrategy)
    {
        Assert.True(AutoMapper.TryCreateClassMap(emptyValueStrategy, out ExcelClassMap<TestClass>? classMap));
        Assert.Equal(emptyValueStrategy, classMap!.EmptyValueStrategy);
        Assert.Equal(typeof(TestClass), classMap.Type);
        Assert.Equal(5, classMap.Properties.Count);

        List<string> members = [.. classMap.Properties.Select(m => m.Member.Name)];
        Assert.Contains("_inheritedField", members);
        Assert.Contains("_field", members);
        Assert.Contains("InheritedProperty", members);
        Assert.Contains("Property", members);
        Assert.Contains("PrivateGetProperty", members);
    }

    [Fact]
    public void TryCreateClass_Map_ObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<object>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_StringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<string>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ByteType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<byte>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SByteType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<sbyte>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_UInt16_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ushort>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_Int16_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<short>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IntType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<int>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_UIntType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<uint>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_UInt64_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<long>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_Int64_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<long>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_FloatType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<float>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_DoubleType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<double>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_DecimalType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<decimal>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_BoolType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<bool>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_DateTimeType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<DateTime>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_MapGuidType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Guid>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_MapEnumType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<MyEnum>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_MapUri_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<MyEnum>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IConvertibleImplementer_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ConvertibleClass>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ArrayStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<string[]>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IEnumerableType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IEnumerable>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IEnumerableObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IEnumerable<object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IEnumerableStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IEnumerable<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IEnumerableGenericType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IEnumerable<IList<string>>>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ICollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ICollection>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ICollectionObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ICollection<object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ICollectionStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ICollection<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IReadOnlyCollectionObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IReadOnlyCollection<object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IList>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IListObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IList<object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IListStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IList<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IReadOnlyListStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IReadOnlyList<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ISetObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ISet<object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ISetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ISet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IReadOnlySetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IReadOnlySet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ArrayListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ArrayList>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_CollectionBaseType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<CollectionBase>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubCollectionBaseType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubCollectionBase>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubCollectionBase : CollectionBase
    {
    }

    [Fact]
    public void TryCreateClass_Map_ReadOnlyCollectionBaseType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ReadOnlyCollectionBase>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubReadOnlyCollectionBaseType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubReadOnlyCollectionBase>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubReadOnlyCollectionBase : ReadOnlyCollectionBase
    {
    }

    [Fact]
    public void TryCreateClass_Map_StackType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Stack>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_QueueType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Queue>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_BitArrayType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<BitArray>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_StringCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<StringCollection>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_CollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Collection<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ReadOnlyCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ReadOnlyCollection<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ObservableCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ObservableCollection<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ReadOnlyObservableCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ReadOnlyObservableCollection<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_LinkedListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<LinkedList<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ListStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<List<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_HashSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<HashSet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ReadOnlySetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ReadOnlySet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SortedSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SortedSet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_FrozenSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<FrozenSet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ImmutableArrayStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ImmutableArray<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ImmutableListStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ImmutableList<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IImmutableListStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IImmutableList<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ImmutableStackStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ImmutableStack<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IImmutableStackStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IImmutableStack<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ImmutableQueueStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ImmutableQueue<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IImmutableQueueStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IImmutableQueue<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ImmutableSortedSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ImmutableSortedSet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ImmutableHashSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ImmutableHashSet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IImmutableSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IImmutableSet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IEnumerableKeyValuePairStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IEnumerable<KeyValuePair<string, object>>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IEnumerableKeyValuePairType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IEnumerable<KeyValuePair<string, string>>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ICollectionKeyValuePairType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ICollection<KeyValuePair<string, string>>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IListKeyValuePairType_ReturnsTrue()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IList<KeyValuePair<string, string>>>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_HashtableType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Hashtable>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_DictionaryBaseType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<DictionaryBase>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubDictionaryBaseType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubDictionaryBase>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubDictionaryBase : DictionaryBase
    {
    }

    [Fact]
    public void TryCreateClass_Map_HybridDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<HybridDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_NameValueCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<NameValueCollection>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ListDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ListDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SortedListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SortedList>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_StringDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<StringDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_OrderedDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<OrderedDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_DictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Dictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_OrderedDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<OrderedDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SortedDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SortedDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_FrozenDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<FrozenDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ReadOnlyDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ReadOnlyDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ConcurrentDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ConcurrentDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SortedListTType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SortedList<string, int>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_KeyedCollectionType_ReturnsTrue()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<KeyedCollection<string, object>>? classMap));
        Assert.Null(classMap);
    }

    private class SubKeyedCollection<TKey, TValue> : KeyedCollection<TKey, TValue> where TKey : notnull
    {
        protected override TKey GetKeyForItem(TValue item) => throw new NotImplementedException();
    }

    [Fact]
    public void TryCreateClass_Map_ImmutableDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ImmutableDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IImmutableDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IImmutableDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ImmutableSortedDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ImmutableSortedDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_CustomIEnumerableKeyValuePairStringObject_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<CustomAddIEnumerable>? classMap));
        Assert.NotNull(classMap);
    }

    private class CustomAddIEnumerable : IEnumerable<KeyValuePair<string, object>>
    {
        public IEnumerator<KeyValuePair<string, object>> GetEnumerator() => throw new NotImplementedException();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    [Fact]
    public void TryCreateClass_Map_ExpandoObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ExpandoObject>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_InterfaceType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IDisposable>? classMap));
        Assert.Null(classMap);
    }

    [Theory]
    [InlineData(FallbackStrategy.ThrowIfPrimitive - 1)]
    [InlineData(FallbackStrategy.SetToDefaultValue + 1)]
    public void TryCreateClass_Map_InvalidFallbackStrategy_ThrowsArgumentException(FallbackStrategy emptyValueStrategy)
    {
        Assert.Throws<ArgumentException>("emptyValueStrategy", () => AutoMapper.TryCreateClassMap<TestClass>(emptyValueStrategy, out _));
    }

    private enum MyEnum
    {
    }

    private class ConvertibleClass : IConvertible
    {
        public TypeCode GetTypeCode() => throw new NotImplementedException();

        public bool ToBoolean(IFormatProvider? provider) => throw new NotImplementedException();

        public byte ToByte(IFormatProvider? provider) => throw new NotImplementedException();

        public char ToChar(IFormatProvider? provider) => throw new NotImplementedException();

        public DateTime ToDateTime(IFormatProvider? provider) => throw new NotImplementedException();

        public decimal ToDecimal(IFormatProvider? provider) => throw new NotImplementedException();

        public double ToDouble(IFormatProvider? provider) => throw new NotImplementedException();

        public short ToInt16(IFormatProvider? provider) => throw new NotImplementedException();

        public int ToInt32(IFormatProvider? provider) => throw new NotImplementedException();

        public long ToInt64(IFormatProvider? provider) => throw new NotImplementedException();

        public sbyte ToSByte(IFormatProvider? provider) => throw new NotImplementedException();

        public float ToSingle(IFormatProvider? provider) => throw new NotImplementedException();

        public string ToString(IFormatProvider? provider) => throw new NotImplementedException();

        public object ToType(Type conversionType, IFormatProvider? provider) => throw new NotImplementedException();

        public ushort ToUInt16(IFormatProvider? provider) => throw new NotImplementedException();

        public uint ToUInt32(IFormatProvider? provider) => throw new NotImplementedException();

        public ulong ToUInt64(IFormatProvider? provider) => throw new NotImplementedException();
    }

#pragma warning disable 0649, 0169
    private class BaseClass
    {
        public string _inheritedField = default!;
        public string InheritedProperty { get; set; } = default!;
    }

    private class TestClass : BaseClass
    {
        public string _field = default!;
        protected string _protectedField = default!;
        internal string _internalField = default!;
#pragma warning disable CS0414 // Field is never assigned to, and will always have its default value
        private readonly string _privateField = default!;
#pragma warning restore CS0414 // Field is never assigned to, and will always have its default value
        public static string s_field = default!;
        
        public string Property { get; set; } = default!;
        protected string ProtectedProperty { get; set; } = default!;
        internal string InternalProperty { get; set; } = default!;
        private string PrivateProperty { get; set; } = default!;
        public static string StaticProperty { get; set; } = default!;

        public string PrivateSetProperty { get; private set; } = default!;
        public string PrivateGetProperty { private get; set; } = default!;
    }
#pragma warning restore 0649
}