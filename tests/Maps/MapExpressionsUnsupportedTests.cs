using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using Xunit;

namespace ExcelMapper.Tests;

public class MapExpressionsUnsupportedTests
{
    [Fact]
    public void Map_NoExpression_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<SimpleClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p));
    }

    [Fact]
    public void Map_CastNoExpression_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<SimpleClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => (object)p));
    }

    [Fact]
    public void Map_NewExpression_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<SimpleClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => new List<string>()));
    }

    [Fact]
    public void Map_MethodCallExpression_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<SimpleClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value.ToString()));
    }

    private class SimpleClass
    {
        public string Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ArrayElementCantBeMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NonConstructibleArrayElementClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => p.Value));
    }

    private class NonConstructibleArrayElementClass
    {
        public IDisposable[] Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ListCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NonConstructibleListClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IList<string>)p.Value));
    }

    private class NonConstructibleListClass
    {
        public NonConstructibleList Value { get; set; } = default!;
    }

    private class NonConstructibleList : List<string>
    {
        private NonConstructibleList() { }
    }

    [Fact]
    public void Map_ImmutableArrayBuilder_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ImmutableArrayBuilderIntClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IList<int>)p.Value));
    }

    public class ImmutableArrayBuilderIntClass
    {
        public ImmutableArray<int>.Builder Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ListElementCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NonConstructibleListElementClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => p.Value));
    }

    private class NonConstructibleListElementClass
    {
        public List<IDisposable> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ImmutableListBuilder_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ImmutableListBuilderIntClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IList<int>)p.Value));
    }

    public class ImmutableListBuilderIntClass
    {
        public ImmutableList<int>.Builder Value { get; set; } = default!;
    }

    [Fact]
    public void Map_DictionaryCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NonConstructibleDictionaryClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IDictionary<string, string>)p.Value));
    }

    private class NonConstructibleDictionaryClass
    {
        public NonConstructibleDictionary Value { get; set; } = default!;
    }

    private class NonConstructibleDictionary : Dictionary<string, string>
    {
        private NonConstructibleDictionary() { }
    }

    [Fact]
    public void Map_ImmutableDictionaryBuilder_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ImmutableDictionaryBuilderIntClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IDictionary<string, int>)p.Value));
    }

    public class ImmutableDictionaryBuilderIntClass
    {
        public ImmutableDictionary<string, int>.Builder Value { get; set; } = default!;
    }

    [Fact]
    public void Map_DictionaryKeyCantBeConstructed_Success()
    {
        var map = new ExcelClassMap<NonConstructibleDictionaryKeyClass>();
        map.Map(p => p.Value);
    }

    private class NonConstructibleDictionaryKeyClass
    {
        public Dictionary<IDisposable, string> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_DictionaryValueCantBeConstructed_Success()
    {
        var map = new ExcelClassMap<NonConstructibleDictionaryValueClass>();
        map.Map(p => p.Value);
    }

    private class NonConstructibleDictionaryValueClass
    {
        public Dictionary<string, IDisposable> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_NonConstantArrayIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ArrayClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[int.Parse("0")]));
    }

    [Fact]
    public void Map_NegativeArrayIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ArrayClass>();
#pragma warning disable CS0251 // Indexing an array with a negative index
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[-1]));
#pragma warning restore CS0251 // Indexing an array with a negative index
    }

    [Fact]
    public void Map_MultidimensionalArrayValue_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<MultidimensionalArrayClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Values[0, 0]));
    }

    private class MultidimensionalArrayClass
    {
        public string[,] Values { get; set; } = default!;
    }

    [Fact]
    public void Map_NonConstantArrayParent_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ArrayClass>();
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value.ToString()[int.Parse("0")]));
#pragma warning restore CS8602 // Dereference of a possibly null reference.
    }

    [Fact]
    public void Map_ArrayIndexerChainedWithoutMember_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ChainedArrayClass>();
        // Chained array access starting with a method call
        Assert.Throws<ArgumentException>(() => map.Map(o => o.GetNestedArray()[0][1]));
    }

    private class ChainedArrayClass
    {
        public int[][] GetNestedArray() => [[1, 2], [3, 4]];
    }

    [Fact]
    public void Map_ArrayWithVariableIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ArrayClass>();
        var index = 0;
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[index]));
    }

    [Fact]
    public void Map_ArrayWithMemberIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ArrayClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[p.IntValue]));
    }

    private class ArrayClass
    {
        public string[] Value { get; set; } = default!;
        public int IntValue { get; set; }
    }

    [Fact]
    public void ReadRow_ArrayIndexDifferently_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<ArrayClass>();
        map.Map(o => o.Value);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0]));
    }

    [Fact]
    public void Map_NonConstantListIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ListClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[int.Parse("0")]));
    }

    [Fact]
    public void Map_NegativeListIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ListClass>();
#pragma warning disable CS0251 // Indexing an array with a negative index
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[-1]));
#pragma warning restore CS0251 // Indexing an array with a negative index
    }

    [Fact]
    public void Map_ListWithVariableIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ListClass>();
        var index = 0;
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[index]));
    }

    [Fact]
    public void Map_ListWithMemberIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ListClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[p.IntValue]));
    }

    private class ListClass
    {
        public List<string> Value { get; set; } = default!;

        public int IntValue { get; set; }
    }

    [Fact]
    public void Map_NonConstantListParent_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<DictionaryClass>();
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value.ToString()[int.Parse("0")]));
#pragma warning restore CS8602 // Dereference of a possibly null reference.
    }

    [Fact]
    public void Map_ListIndexerElementImmutableArrayBuilderInt_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ImmutableArrayBuilderIntClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => p.Value[0]));
    }

    [Fact]
    public void ReadRow_ListIndexDifferently_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<ListClass>();
        map.Map(o => o.Value);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0]));
    }

    [Fact]
    public void Map_NonConstantDictionaryKey_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<DictionaryClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value["key".ToString()]));
    }

    [Fact]
    public void Map_NonStringDictionaryKey_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<IntDictionaryClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => p.Value[0]));
    }

    private class IntDictionaryClass
    {
        public Dictionary<int, string> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_NonConstantDictionaryParent_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<DictionaryClass>();
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.GetDictionary()["key"]));
#pragma warning restore CS8602 // Dereference of a possibly null reference.
    }

    [Fact]
    public void Map_DictionaryWithNullKey_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<DictionaryClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[null!]));
    }

    [Fact]
    public void Map_DictionaryWithVariableKey_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<DictionaryClass>();
        var key = "key";
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[key]));
    }

    [Fact]
    public void Map_DictionaryWithMemberKey_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<DictionaryClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[p.StringValue]));
    }

    private class DictionaryClass
    {
        public Dictionary<string, string> Value { get; set; } = default!;
        public Dictionary<string, string> GetDictionary() => Value;
        public string StringValue { get; set; } = default!;
    }


    [Fact]
    public void Map_DictionaryIndexerCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NonConstructibleDictionaryClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => p.Value["key"]));
    }

    [Fact]
    public void ReadRow_DictionaryIndexDifferently_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<DictionaryClass>();
        map.Map(o => o.Value);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value["key"]));
    }
}