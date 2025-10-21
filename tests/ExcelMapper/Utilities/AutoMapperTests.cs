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
using System.Numerics;
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
    public void TryCreateClass_Map_UIntPtr_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<UIntPtr>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IntPtr_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IntPtr>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_UInt128_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<UInt128>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_Int128_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Int128>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_BigInteger_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<BigInteger>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_Complex_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Complex>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_HalfType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Half>? classMap));
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
    public void TryCreateClass_Map_DateTimeOffsetType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<DateTimeOffset>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_TimeSpanType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<TimeSpan>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_DateOnlyType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<DateOnly>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_TimeOnlyType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<TimeOnly>? classMap));
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
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Enum>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_MapCustomEnumType_ReturnsTrue()
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
    public void TryCreateClass_Map_VersionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Version>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IConvertibleImplementer_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ConvertibleClass>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ArrayType_ReturnsFalse()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Array>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ArrayStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<string[]>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_MultidimensionalArrayStringType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<string[,]>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IEnumerableType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IEnumerable>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubIEnumerableType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIEnumerable>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIEnumerable : IEnumerable
    {
    }

    [Fact]
    public void TryCreateClass_Map_IEnumerableObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IEnumerable<object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubIEnumerableObjectType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIEnumerableObject>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIEnumerableObject : IEnumerable<object>
    {
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
    public void TryCreateClass_Map_SubICollectionType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubICollection>? classMap));
        Assert.Null(classMap);
    }

    private interface SubICollection : ICollection
    {
    }

    [Fact]
    public void TryCreateClass_Map_ICollectionObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ICollection<object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubICollectionObjectType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubICollectionObject>? classMap));
        Assert.Null(classMap);
    }

    private interface SubICollectionObject : ICollection<object>
    {
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
    public void TryCreateClass_Map_SubIReadOnlyCollectionObjectType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIReadOnlyCollectionObject>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIReadOnlyCollectionObject : IReadOnlyCollection<object>
    {
    }

    [Fact]
    public void TryCreateClass_Map_IListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IList>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubIListType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIList>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIList : IList
    {
    }

    [Fact]
    public void TryCreateClass_Map_IListObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IList<object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubIListObjectType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIListObject>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIListObject : IList<object>
    {
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
    public void TryCreateClass_Map_SubISetObjectType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubISetObject>? classMap));
        Assert.Null(classMap);
    }

    private interface SubISetObject : ISet<object>
    {
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
    public void TryCreateClass_Map_SubIReadOnlySetStringType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIReadOnlySetString>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIReadOnlySetString : IReadOnlySet<string>
    {
    }

    [Fact]
    public void TryCreateClass_Map_ArrayListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ArrayList>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubArrayListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubArrayList>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubArrayList : ArrayList
    {
    }

    [Fact]
    public void TryCreateClass_Map_BitArrayType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<BitArray>? classMap));
        Assert.Null(classMap);
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
    public void TryCreateClass_Map_QueueType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Queue>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubQueueType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubQueue>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubQueue : Queue
    {
    }

    [Fact]
    public void TryCreateClass_Map_SortedListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SortedList>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubSortedListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubSortedList>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubSortedList : SortedList
    {
    }

    [Fact]
    public void TryCreateClass_Map_StackType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Stack>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubStackType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubStack>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubStack : Stack
    {
    }

    [Fact]
    public void TryCreateClass_Map_BitVector32Type_ReturnsTrue()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<BitVector32>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_StringCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<StringCollection>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubStringCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubStringCollection>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubStringCollection : StringCollection
    {
    }

    [Fact]
    public void TryCreateClass_Map_CollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Collection<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubCollection>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubCollection : Collection<string>
    {
    }

    [Fact]
    public void TryCreateClass_Map_KeyedCollectionType_ReturnsTrue()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<KeyedCollection<string, object>>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubKeyedCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubKeyedCollection<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubKeyedCollection<TKey, TValue> : KeyedCollection<TKey, TValue> where TKey : notnull
    {
        protected override TKey GetKeyForItem(TValue item) => throw new NotImplementedException();
    }

    [Fact]
    public void TryCreateClass_Map_ReadOnlyCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ReadOnlyCollection<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubReadOnlyCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubReadOnlyCollection>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubReadOnlyCollection : ReadOnlyCollection<string>
    {
        public SubReadOnlyCollection(IList<string> list) : base(list)
        {
        }
    }

    [Fact]
    public void TryCreateClass_Map_ObservableCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ObservableCollection<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubObservableCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubObservableCollection>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubObservableCollection : ObservableCollection<string>
    {
    }

    [Fact]
    public void TryCreateClass_Map_ReadOnlyObservableCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ReadOnlyObservableCollection<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubReadOnlyObservableCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubReadOnlyObservableCollection>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubReadOnlyObservableCollection : ReadOnlyObservableCollection<string>
    {
        public SubReadOnlyObservableCollection(ObservableCollection<string> list) : base(list)
        {
        }
    }

    [Fact]
    public void TryCreateClass_Map_LinkedListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<LinkedList<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubLinkedListType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubLinkedList>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubLinkedList : LinkedList<string>
    {
    }

    [Fact]
    public void TryCreateClass_Map_ListStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<List<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubListStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubList>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubList : List<string>
    {
    }

    [Fact]
    public void TryCreateClass_Map_HashSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<HashSet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubHashSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubHashSet>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubHashSet : HashSet<string>
    {
    }

    [Fact]
    public void TryCreateClass_Map_ReadOnlySetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ReadOnlySet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubReadOnlySetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubReadOnlySet>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubReadOnlySet : ReadOnlySet<string>
    {
        public SubReadOnlySet(ISet<string> set) : base(set)
        {
        }
    }

    [Fact]
    public void TryCreateClass_Map_SortedSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SortedSet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubSortedSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubSortedSet>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubSortedSet : SortedSet<string>
    {
        public SubSortedSet() : base()
        {
        }

        public SubSortedSet(IComparer<string> comparer) : base(comparer)
        {
        }

        public SubSortedSet(IEnumerable<string> collection) : base(collection)
        {
        }

        public SubSortedSet(IEnumerable<string> collection, IComparer<string> comparer) : base(collection, comparer)
        {
        }
    }

    [Fact]
    public void TryCreateClass_Map_BlockingCollectionStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<BlockingCollection<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ConcurrentBagStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ConcurrentBag<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ConcurrentQueueStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ConcurrentQueue<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_ConcurrentStackStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ConcurrentStack<string>>? classMap));
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
    public void TryCreateClass_Map_SubIImmutableListStringType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIImmutableList>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIImmutableList : IImmutableList<string>
    {
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
    public void TryCreateClass_Map_SubIImmutableStackStringType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIImmutableStack>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIImmutableStack : IImmutableStack<string>
    {
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
    public void TryCreateClass_Map_SubIImmutableQueueStringType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIImmutableQueue>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIImmutableQueue : IImmutableQueue<string>
    {
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
    public void TryCreateClass_Map_SubIImmutableSetStringType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIImmutableSet>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIImmutableSet : IImmutableSet<string>
    {
    }

    [Fact]
    public void TryCreateClass_Map_FrozenSetStringType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<FrozenSet<string>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_IEnumerableKeyValuePairStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IEnumerable<KeyValuePair<string, object>>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubIEnumerableKeyValuePairStringObjectType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIEnumerableKeyValuePairStringObject>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIEnumerableKeyValuePairStringObject : IEnumerable<KeyValuePair<string, object>>
    {
    }

    [Fact]
    public void TryCreateClass_Map_IEnumerableKeyValuePairType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IEnumerable<KeyValuePair<string, string>>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubIEnumerableKeyValuePairType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIEnumerableKeyValuePair>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIEnumerableKeyValuePair : IEnumerable<KeyValuePair<string, string>>
    {
    }

    [Fact]
    public void TryCreateClass_Map_ICollectionKeyValuePairType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ICollection<KeyValuePair<string, string>>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubICollectionKeyValuePairType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubICollectionKeyValuePair>? classMap));
        Assert.Null(classMap);
    }

    private interface SubICollectionKeyValuePair : ICollection<KeyValuePair<string, string>>
    {
    }

    [Fact]
    public void TryCreateClass_Map_IListKeyValuePairType_ReturnsTrue()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IList<KeyValuePair<string, string>>>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubIListKeyValuePairType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIListKeyValuePair>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIListKeyValuePair : IList<KeyValuePair<string, string>>
    {
    }

    [Fact]
    public void TryCreateClass_Map_IDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubIDictionaryType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIDictionary>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIDictionary : IDictionary
    {
    }

    [Fact]
    public void TryCreateClass_Map_IDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<IDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubIDictionaryStringObjectType_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubIDictionaryStringObject>? classMap));
        Assert.Null(classMap);
    }

    private interface SubIDictionaryStringObject : IDictionary<string, object>
    {
    }

    [Fact]
    public void TryCreateClass_Map_HashtableType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Hashtable>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubHashtableType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubHashtable>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubHashtable : Hashtable
    {
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
    public void TryCreateClass_Map_SubHybridDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubHybridDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubHybridDictionary : HybridDictionary
    {
    }

    [Fact]
    public void TryCreateClass_Map_ListDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ListDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubListDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubListDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubListDictionary : ListDictionary
    {
    }

    [Fact]
    public void TryCreateClass_Map_NameObjectCollectionBase_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<NameObjectCollectionBase>? classMap));
        Assert.Null(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubNameObjectCollectionBase_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubNameObjectCollectionBase>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubNameObjectCollectionBase : NameObjectCollectionBase
    {
    }

    [Fact]
    public void TryCreateClass_Map_NameValueCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<NameValueCollection>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubNameValueCollectionType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubNameValueCollection>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubNameValueCollection : NameValueCollection
    {
    }

    [Fact]
    public void TryCreateClass_Map_StringDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<StringDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubStringDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubStringDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubStringDictionary : StringDictionary
    {
    }

    [Fact]
    public void TryCreateClass_Map_OrderedDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<OrderedDictionary>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubOrderedDictionaryType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubOrderedDictionary>? classMap));
        Assert.NotNull(classMap);
    }
    
    private class SubOrderedDictionary : OrderedDictionary
    {
    }

    [Fact]
    public void TryCreateClass_Map_DictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<Dictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }
    
    [Fact]
    public void TryCreateClass_Map_SubDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubDictionaryStringObject>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubDictionaryStringObject : Dictionary<string, object>
    {
    }

    [Fact]
    public void TryCreateClass_Map_OrderedDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<OrderedDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubOrderedDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubOrderedDictionaryStringObject>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubOrderedDictionaryStringObject : OrderedDictionary<string, object>
    {
    }

    [Fact]
    public void TryCreateClass_Map_SortedDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SortedDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubSortedDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubSortedDictionaryStringObject>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubSortedDictionaryStringObject : SortedDictionary<string, object>
    {
    }

    [Fact]
    public void TryCreateClass_Map_ReadOnlyDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ReadOnlyDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubReadOnlyDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubReadOnlyDictionaryStringObject>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubReadOnlyDictionaryStringObject : ReadOnlyDictionary<string, object>
    {
        public SubReadOnlyDictionaryStringObject(IDictionary<string, object> dictionary) : base(dictionary)
        {
        }
    }

    [Fact]
    public void TryCreateClass_Map_ConcurrentDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<ConcurrentDictionary<string, object>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubConcurrentDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubConcurrentDictionaryStringObject>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubConcurrentDictionaryStringObject : ConcurrentDictionary<string, object>
    {
    }

    [Fact]
    public void TryCreateClass_Map_SortedListTType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SortedList<string, int>>? classMap));
        Assert.NotNull(classMap);
    }

    [Fact]
    public void TryCreateClass_Map_SubSortedListTType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<SubSortedListT>? classMap));
        Assert.NotNull(classMap);
    }

    private class SubSortedListT : SortedList<string, int>
    {
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
    public void TryCreateClass_Map_FrozenDictionaryStringObjectType_ReturnsTrue()
    {
        Assert.True(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<FrozenDictionary<string, object>>? classMap));
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

    [Fact]
    public void TryCreateClass_Map_InterfaceTypeProperty_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<InterfaceProperty>? classMap));
        Assert.Null(classMap);
    }

    private class InterfaceProperty
    {
        public IDisposable Property { get; set; } = default!;
    }

    [Fact]
    public void TryCreateClass_Map_InterfaceTypeField_ReturnsFalse()
    {
        Assert.False(AutoMapper.TryCreateClassMap(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<InterfaceField>? classMap));
        Assert.Null(classMap);
    }

    private class InterfaceField
    {
        public IDisposable Field = default!;
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