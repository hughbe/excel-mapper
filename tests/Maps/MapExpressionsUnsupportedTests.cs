using System.Collections.Immutable;

namespace ExcelMapper.Tests;

public class MapExpressionsUnsupportedTests
{
    [Fact]
    public void Map_NoExpression_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p));
    }

    [Fact]
    public void Map_CastNoExpression_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => (object)p));
    }

    [Fact]
    public void Map_NewExpression_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => new List<string>()));
    }

    [Fact]
    public void Map_MethodCallExpression_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value.ToString()));
    }

    [Fact]
    public void Map_MemberVariable_ThrowsArgumentException()
    {
        var otherType = new StringClass();
        var map = new ExcelClassMap<StringClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => otherType));
    }

    [Fact]
    public void Map_MemberInvalidTargetType_ThrowsArgumentException()
    {
        var otherType = new StringClass();
        var map = new ExcelClassMap<StringClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => otherType.Value));
    }

    [Fact]
    public void Map_InvalidUnaryExpression_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<IntClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => -p.Value));
    }

    [Fact]
    public void Map_InvalidBinaryExpressionFirst_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<IntClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => 1 + p.Value));
    }

    [Fact]
    public void Map_InvalidBinaryExpressionSecond_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<IntClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value + 1));
    }

    private class StringClass
    {
        public string Value { get; set; } = default!;
    }

    private class IntClass
    {
        public int Value { get; set; } = default!;
    }
    [Fact]
    public void Map_StaticProperty_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StaticMemberClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => StaticMemberClass.StaticValue));
    }

    [Fact]
    public void Map_StaticField_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StaticMemberClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => StaticMemberClass.StaticField));
    }

    [Fact]
    public void Map_StaticPropertyInChain_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ChainedStaticClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => ChainedStaticClass.Inner.Value));
    }

    private class StaticMemberClass
    {
        public static int StaticValue { get; set; }
#pragma warning disable CS0649 // Field is never assigned to, and will always have its default value
        public static int StaticField;
#pragma warning restore CS0649 // Field is never assigned to, and will always have its default value
    }

    private class ChainedStaticClass
    {
        public static InnerStaticClass Inner { get; set; } = new InnerStaticClass();
    }

    private class InnerStaticClass
    {
        public int Value { get; set; }
    }

    [Fact]
    public void Map_NestedClassDifferently_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NestedClassParent>();
        map.Map(o => o.Value);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value.Value));
    }

    private class NestedClassParent
    {
        public NestedClassChild Value { get; set; } = default!;
    }

    private class NestedClassChild
    {
        public int Value { get; set; }
    }

    [Fact]
    public void Map_ArrayElementCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NonDisposableArrayElementClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => p.Value));
        Assert.Throws<ExcelMappingException>(() => map.MapList<NonConstructibleArrayElement>(p => p.Value));
    }

    private class NonDisposableArrayElementClass
    {
        public NonConstructibleArrayElement[] Value { get; set; } = default!;
    }

    private class NonConstructibleArrayElement
    {
        private NonConstructibleArrayElement() { }
    }

    [Fact]
    public void Map_CastArrayElementCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ObjectClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (NonConstructibleArrayElement[])p.Value));
        Assert.Throws<ExcelMappingException>(() => map.MapList<NonConstructibleArrayElement>(p => (NonConstructibleArrayElement[])p.Value));
    }

    [Fact]
    public void Map_ArrayElementCantBeMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<IDisposableArrayElementClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => p.Value));
        Assert.Throws<ExcelMappingException>(() => map.MapList<IDisposable>(p => p.Value));
    }

    private class IDisposableArrayElementClass
    {
        public IDisposable[] Value { get; set; } = default!;
    }

    [Fact]
    public void Map_CastArrayElementCantBeMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ObjectClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IDisposable[])p.Value));
        Assert.Throws<ExcelMappingException>(() => map.MapList<IDisposable>(p => (IDisposable[])p.Value));
    }

    [Fact]
    public void Map_MultidimensionalArrayValue_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<StringMultidimensionalArrayClass>();
        Assert.Throws<ExcelMappingException>(() => map.MapList<string>(p => p.Value));
    }

    [Fact]
    public void Map_MultidimensionalArrayElementCantBeMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<IDisposableStringMultidimensionalArrayClass>();
        Assert.Throws<ExcelMappingException>(() => map.MapList<IDisposable>(p => p.Value));
    }

    private class IDisposableStringMultidimensionalArrayClass
    {
        public IDisposable[,] Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ListCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NonConstructibleListIndexClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IList<string>)p.Value));
        Assert.Throws<ExcelMappingException>(() => map.MapList<string>(p => p.Value));
    }

    [Fact]
    public void Map_CastListCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ObjectClass>();
        Assert.Throws<ExcelMappingException>(() => map.MapList<string>(p => (NonConstructibleList)p.Value));
    }

    private class NonConstructibleListIndexClass
    {
        public NonConstructibleList Value { get; set; } = default!;
    }

    private class NonConstructibleList : List<string>
    {
        private NonConstructibleList() { }
    }

    [Fact]
    public void Map_ListElementCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NonConstructibleListElementClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IList<NonConstructibleListElement>)p.Value));
        Assert.Throws<ExcelMappingException>(() => map.MapList<NonConstructibleListElement>(p => p.Value));
    }

    private class NonConstructibleListElementClass
    {
        public List<NonConstructibleListElement> Value { get; set; } = default!;
    }

    private class NonConstructibleListElement
    {
        private NonConstructibleListElement() { }
    }

    [Fact]
    public void Map_CastListElementCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ObjectClass>();
        Assert.Throws<ExcelMappingException>(() => map.MapList<NonConstructibleListElement>(p => (List<NonConstructibleListElement>)p.Value));
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
    public void Map_ListElementCantBeMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<IDisposableListElementClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IList<IDisposable>)p.Value));
        Assert.Throws<ExcelMappingException>(() => map.MapList<IDisposable>(p => p.Value));
    }

    private class IDisposableListElementClass
    {
        public List<IDisposable> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_CastListElementCantBeMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ObjectClass>();
        Assert.Throws<ExcelMappingException>(() => map.MapList<IDisposable>(p => (List<IDisposable>)p.Value));
    }

    [Fact]
    public void Map_DictionaryCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NonConstructibleDictionaryClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IDictionary<string, string>)p.Value));
        Assert.Throws<ExcelMappingException>(() => map.MapDictionary<string>(p => p.Value));
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
    public void Map_CastDictionaryCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ObjectClass>();
        Assert.Throws<ExcelMappingException>(() => map.MapDictionary<string>(p => (NonConstructibleDictionary)p.Value));
    }

    [Fact]
    public void Map_ImmutableDictionaryBuilder_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ImmutableDictionaryBuilderIntClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => (IDictionary<string, int>)p.Value));
        Assert.Throws<ExcelMappingException>(() => map.MapDictionary<int>(p => p.Value));
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
        map.MapDictionary<string>(p => p.Value);
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
        map.MapDictionary<IDisposable>(p => p.Value);
    }

    private class NonConstructibleDictionaryValueClass
    {
        public Dictionary<string, IDisposable> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_RootDefaultMappedIntArrayIndex_ThrowsArgumentException()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<int[]>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o[0]));
    }

    [Fact]
    public void Map_RootDefaultMappedIntListIndex_ThrowsArgumentException()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<List<int>>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o[0]));
    }

    [Fact]
    public void Map_RootDefaultMappedIntMultidimensionalIndex_ThrowsArgumentException()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<int[,]>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o[0, 0]));
    }

    [Fact]
    public void Map_RootDefaultMappedIntDictionaryIndex_ThrowsArgumentException()
    {
        using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
        var map = new ExcelClassMap<Dictionary<string, int>>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o["key"]));
    }

    [Fact]
    public void Map_StaticArrayIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StaticArrayClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => StaticArrayClass.Value[0]));
    }

    private class StaticArrayClass
    {
        public static string[] Value { get; set; } = default!;
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
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.GetNestedArray()[0][1]));
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

    [Fact]
    public void Map_ArrayIndexMemberDifferently_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<ArrayClass>();
        map.Map(o => o.Value);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0]));
    }

    private class ArrayClass
    {
        public string[] Value { get; set; } = default!;
        public int IntValue { get; set; }
    }

    [Fact]
    public void Map_ArrayIndexDifferentlyArrayIndex_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<ArrayArrayClass>();
        map.Map(o => o.Value[0]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0][0]));
    }

    private class ArrayArrayClass
    {
        public int[][] Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ArrayIndexDifferentlyMultidimensionalIndex_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<IntMultidimensionalArrayClass>();
        map.Map(o => o.Value[0, 0]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0, 0][0]));
    }

    private class IntMultidimensionalArrayClass
    {
        public int[,][] Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ArrayIndexDifferentlyListIndex_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<ListArrayClass>();
        map.Map(o => o.Value[0]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0][0]));
    }

    private class ListArrayClass
    {
        public List<int[]> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ArrayIndexDifferentlyDictionaryKey_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<DictionaryArrayClass>();
        map.Map(o => o.Value["key"]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value["key"][0]));
    }

    private class DictionaryArrayClass
    {
        public Dictionary<string, int[]> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_StaticMultidimensionalIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StaticStringMultidimensionalArrayClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => StaticStringMultidimensionalArrayClass.Value[0, 0]));
    }

    private class StaticStringMultidimensionalArrayClass
    {
        public static string[,] Value { get; set; } = default!;
    }

    [Fact]
    public void Map_NonConstantMiltidimensionalIndexFirst_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringMultidimensionalArrayClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[int.Parse("0"), 0]));
    }

    [Fact]
    public void Map_NonConstantMiltidimensionalIndexSecond_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringMultidimensionalArrayClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[0, int.Parse("0")]));
    }

    [Fact]
    public void Map_NegativeMultidimensionalIndexFirst_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringMultidimensionalArrayClass>();
#pragma warning disable CS0251 // Indexing an array with a negative index
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[-1, 0]));
#pragma warning restore CS0251 // Indexing an array with a negative index
    }

    [Fact]
    public void Map_NegativeMultidimensionalIndexSecond_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringMultidimensionalArrayClass>();
#pragma warning disable CS0251 // Indexing an array with a negative index
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[0, -1]));
#pragma warning restore CS0251 // Indexing an array with a negative index
    }

    [Fact]
    public void Map_MultidimensionalIndexerChainedWithoutMember_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ChainedStringMultidimensionalArrayClass>();
        // Chained array access starting with a method call
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.GetNestedArray()[0][1]));
    }

    private class ChainedStringMultidimensionalArrayClass
    {
        public int[][] GetNestedArray() => [[1, 2], [3, 4]];
    }

    [Fact]
    public void Map_MultidimensionalWithVariableIndexFirst_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringMultidimensionalArrayClass>();
        var index = 0;
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[index, 0]));
    }

    [Fact]
    public void Map_MultidimensionalWithVariableIndexSecond_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringMultidimensionalArrayClass>();
        var index = 0;
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[0, index]));
    }

    [Fact]
    public void Map_MultidimensionalWithMemberIndexFirst_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringMultidimensionalArrayClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[p.IntValue, 0]));
    }

    [Fact]
    public void Map_MultidimensionalWithMemberIndexSecond_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StringMultidimensionalArrayClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[0, p.IntValue]));
    }

    private class StringMultidimensionalArrayClass
    {
        public string[,] Value { get; set; } = default!;
        public int IntValue { get; set; }
    }

    [Fact]
    public void Map_MultidimensionalClassHasGetNoArguments_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClassHasGetNoArgumentsClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.Value.Get()));
    }


    private class MultidimensionalClassHasGetNoArgumentsClass
    {
        public MultidimensionalClassHasGetNoArguments Value { get; set; } = default!;
    }

    private class MultidimensionalClassHasGetNoArguments
    {
        public int Get() => 0;
    }

    [Fact]
    public void Map_MultidimensionalClassHasGetOneArgument_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClassHasGetOneArgumentClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.Value.Get(0)));
    }

    private class MultidimensionalClassHasGetOneArgumentClass
    {
        public MultidimensionalClassHasGetOneArgument Value { get; set; } = default!;
    }

    private class MultidimensionalClassHasGetOneArgument
    {
        public int Get(int i) => 0;
    }

    [Fact]
    public void Map_MultidimensionalClassHasGetTwoArguments_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClassHasGetTwoArgumentsClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.Value.Get(0, 1)));
    }

    private class MultidimensionalClassHasGetTwoArgumentsClass
    {
        public MultidimensionalClassHasGetTwoArguments Value { get; set; } = default!;
    }

    private class MultidimensionalClassHasGetTwoArguments
    {
        public int Get(int i, int j) => 0;
    }

    [Fact]
    public void Map_MultidimensionalClassHasGetNonIntegerArgumentFirst_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClassHasGetTwoArgumentsNonIntegerFirstClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.Value.Get("0", 1)));
    }

    private class MultidimensionalClassHasGetTwoArgumentsNonIntegerFirstClass
    {
        public MultidimensionalClassHasGetTwoArgumentsNonIntegerFirst Value { get; set; } = default!;
    }

    private class MultidimensionalClassHasGetTwoArgumentsNonIntegerFirst
    {
        public int Get(string i, int j) => 0;
    }

    [Fact]
    public void Map_MultidimensionalClassHasGetNonIntegerArgumentSecond_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClassHasGetTwoArgumentsNonIntegerSecondClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.Value.Get(0, "0")));
    }

    private class MultidimensionalClassHasGetTwoArgumentsNonIntegerSecondClass
    {
        public MultidimensionalClassHasGetTwoArgumentsNonIntegerSecond Value { get; set; } = default!;
    }

    private class MultidimensionalClassHasGetTwoArgumentsNonIntegerSecond
    {
        public int Get(int i, string j) => 0;
    }

    [Fact]
    public void Map_MultidimensionalClassHasGetNonConstantArgumentFirst_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClassHasGetTwoArgumentsNonConstantFirstClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.Value.Get(int.Parse("0"), 0)));
    }

    private class MultidimensionalClassHasGetTwoArgumentsNonConstantFirstClass
    {
        public MultidimensionalClassHasGetTwoArgumentsNonConstantFirst Value { get; set; } = default!;
    }

    private class MultidimensionalClassHasGetTwoArgumentsNonConstantFirst
    {
        public int Get(int i, int j) => 0;
    }

    [Fact]
    public void Map_MultidimensionalClassHasGetNonConstantArgumentSecond_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClassHasGetTwoArgumentsNonConstantSecondClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.Value.Get(0, int.Parse("0"))));
    }

    private class MultidimensionalClassHasGetTwoArgumentsNonConstantSecondClass
    {
        public MultidimensionalClassHasGetTwoArgumentsNonConstantSecond Value { get; set; } = default!;
    }

    private class MultidimensionalClassHasGetTwoArgumentsNonConstantSecond
    {
        public int Get(int i, int j) => 0;
    }

    [Fact]
    public void Map_MultidimensionalClassHasGetVoidReturnType_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClassHasGetTwoArgumentsNonConstantSecondClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.Value.Get(0, int.Parse("0"))));
    }

    private class MultidimensionalClassHasGetVoidReturnTypeClass
    {
        public MultidimensionalClassHasGetVoidReturnType Value { get; set; } = default!;
    }

    private class MultidimensionalClassHasGetVoidReturnType
    {
        public void Get(int i, int j) { }
    }

    [Fact]
    public void Map_MultidimensionalClassHasGetInvalidReturnType_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClassHasGetInvalidReturnTypeClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => o.Value.Get(0, int.Parse("0"))));
    }

    private class MultidimensionalClassHasGetInvalidReturnTypeClass
    {
        public MultidimensionalClassHasGetInvalidReturnType Value { get; set; } = default!;
    }

    private class MultidimensionalClassHasGetInvalidReturnType
    {
        public string Get(int i, int j) => string.Empty;
    }

    [Fact]
    public void Map_MultidimensionalClassHasGetStatic_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<MultidimensionalClassHasGetStaticClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => MultidimensionalClassHasGetStatic.Get(0, 0)));
    }

    private class MultidimensionalClassHasGetStaticClass
    {
        public MultidimensionalClassHasGetStatic Value { get; set; } = default!;
    }

    private class MultidimensionalClassHasGetStatic
    {
        public static int Get(int i, int j) => 0;
    }

    [Fact]
    public void Map_MultidimensionalIndexDifferently_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<StringMultidimensionalArrayClass>();
        map.Map(o => o.Value);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0, 0]));
    }

    [Fact]
    public void Map_StaticListIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StaticListClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => StaticListClass.Value[0]));
    }

    private class StaticListClass
    {
        public static List<string> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_NonConstantListIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ListIndexClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[int.Parse("0")]));
    }

    [Fact]
    public void Map_NegativeListIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ListIndexClass>();
#pragma warning disable CS0251 // Indexing an array with a negative index
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[-1]));
#pragma warning restore CS0251 // Indexing an array with a negative index
    }

    [Fact]
    public void Map_ListWithVariableIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ListIndexClass>();
        var index = 0;
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[index]));
    }

    [Fact]
    public void Map_ListWithMemberIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ListIndexClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[p.IntValue]));
    }

    private class ListIndexClass
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
    public void Map_ListIndexNonIntArgument_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ListIndexClassGetItemNonIntArgumentClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[1.0]));
    }

    private class ListIndexClassGetItemNonIntArgumentClass
    {
        public ListIndexClassGetItemNonIntArgument Value { get; set; } = default!;
    }

    private class ListIndexClassGetItemNonIntArgument
    {
        public int this[double index] => 0;
    }

    [Fact]
    public void Map_ListIndexIndexerNotEnumerable_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ListIndexClassNotEnumerableClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value[1]));
    }

    private class ListIndexClassNotEnumerableClass
    {
        public ListIndexClassNotEnumerable Value { get; set; } = default!;
    }

    private class ListIndexClassNotEnumerable
    {
        public int this[int index] => 0;
    }

    [Fact]
    public void Map_ListIndexerClassHasGetStatic_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ListIndexerClassHasGetStaticClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => ListIndexerClassHasGetStatic.get_Item(0)));
    }

    private class ListIndexerClassHasGetStaticClass
    {
        public ListIndexerClassHasGetStatic Value { get; set; } = default!;
    }

    private class ListIndexerClassHasGetStatic
    {
        public static int get_Item(int i) => 0;
    }

    [Fact]
    public void Map_ListIndexDifferentlyMember_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ListIndexClass>();
        map.Map(o => o.Value);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0]));
    }

    [Fact]
    public void Map_ListIndexDifferentlyArrayIndex_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<ArrayListClass>();
        map.Map(o => o.Value[0]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0][0]));
    }

    private class ArrayListClass
    {
        public List<int[]> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ListIndexDifferentlyMultidimensionalIndex_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<ListMultidimensionalArrayClass>();
        map.Map(o => o.Value[0, 0]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0, 0][0]));
    }

    private class ListMultidimensionalArrayClass
    {
        public List<int>[,] Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ListIndexDifferentlyListIndex_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<ListListClass>();
        map.Map(o => o.Value[0]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0][0]));
    }

    private class ListListClass
    {
        public List<List<int>> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_ListIndexDifferentlyDictionaryKey_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<DictionaryListClass>();
        map.Map(o => o.Value["key"]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value["key"][0]));
    }

    private class DictionaryListClass
    {
        public Dictionary<string, List<int>> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_StaticDictionaryIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<StaticDictionaryClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => StaticDictionaryClass.Value["key"]));
    }

    private class StaticDictionaryClass
    {
        public static Dictionary<string, string> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_NonConstantDictionaryKey_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<DictionaryClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value["key".ToString()]));
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
    public void Map_CastDictionaryIndexerCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<ObjectClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => ((NonConstructibleDictionary)p.Value)["key"]));
    }

    [Fact]
    public void Map_CastInvalidDictionaryIndexerCantBeConstructed_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<DictionaryClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(o => ((NonConstructibleDictionary)o.Value)["key"]));
    }

    [Fact]
    public void Map_DictionaryIndexIndexerNotEnumerable_ThrowArgumentException()
    {
        var map = new ExcelClassMap<DictionaryIndexClassNotEnumerableClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => p.Value["key"]));
    }

    [Fact]
    public void Map_CastDictionaryIndexerNotEnumerable_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<ObjectClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => ((DictionaryIndexClassNotEnumerable)p.Value)["key"]));
    }

    private class DictionaryIndexClassNotEnumerableClass
    {
        public DictionaryIndexClassNotEnumerable Value { get; set; } = default!;
    }

    private class DictionaryIndexClassNotEnumerable
    {
        public int this[string index] => 0;
    }

    [Fact]
    public void Map_DictionaryIndexerClassHasGetStatic_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<DictionaryIndexerClassHasGetStaticClass>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(o => DictionaryIndexerClassHasGetStatic.get_Item("key")));
    }

    private class DictionaryIndexerClassHasGetStaticClass
    {
        public DictionaryIndexerClassHasGetStatic Value { get; set; } = default!;
    }

    private class DictionaryIndexerClassHasGetStatic
    {
        public static int get_Item(string key) => 0;
    }

    [Fact]
    public void Map_DictionaryIndexDifferentlyMember_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<DictionaryClass>();
        map.Map(o => o.Value);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value["key"]));
    }

    [Fact]
    public void Map_DictionaryIndexDifferentlyArrayIndex_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<ArrayDictionaryClass>();
        map.Map(o => o.Value[0]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0]["key"]));
    }

    private class ArrayDictionaryClass
    {
        public Dictionary<string, int>[] Value { get; set; } = default!;
    }

    [Fact]
    public void Map_DictionaryIndexDifferentlyMultidimensionalIndex_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<DictionaryMultidimensionalArrayClass>();
        map.Map(o => o.Value[0, 0]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0, 0]["key"]));
    }

    private class DictionaryMultidimensionalArrayClass
    {
        public Dictionary<string, int>[,] Value { get; set; } = default!;
    }

    [Fact]
    public void Map_DictionaryIndexDifferentlyListIndex_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<ListDictionaryClass>();
        map.Map(o => o.Value[0]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value[0]["key"]));
    }

    private class ListDictionaryClass
    {
        public List<Dictionary<string, int>> Value { get; set; } = default!;
    }

    [Fact]
    public void Map_DictionaryIndexDifferentlyDictionaryKey_ThrowsInvalidOperationException()
    {
        var map = new ExcelClassMap<DictionaryDictionaryClass>();
        map.Map(o => o.Value["key"]);
        Assert.Throws<InvalidOperationException>(() => map.Map(o => o.Value["key"]["key"]));
    }

    private class DictionaryDictionaryClass
    {
        public Dictionary<string, Dictionary<string, int>> Value { get; set; } = default!;
    }

    private class ObjectClass
    {
        public object Value { get; set; } = default!;
    }
}
