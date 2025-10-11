using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using Xunit;

namespace ExcelMapper.Tests;

public class MultiMapTests
{
    [Fact]
    public void ReadRow_DefaultArray_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<string[]>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultIEnumerable_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerable>();
        Assert.Equal(new string[] { "1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" }, row1);
    }
    
    [Fact]
    public void ReadRow_DefaultIEnumerableString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IEnumerable<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3"], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultICollection_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollection>();
        Assert.Equal(new string[] { "1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" }, row1);
    }
    
    [Fact]
    public void ReadRow_DefaultICollectionString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ICollection<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3"], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultIReadOnlyCollectionString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyCollection<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3"], row1);
    }

    [Fact]
    public void ReadRow_DefaultIList_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IList>();
        Assert.Equal(new string[] { "1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" }, row1);
    }
    
    [Fact]
    public void ReadRow_DefaultIListString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IList<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultIReadOnlyListString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<IReadOnlyList<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultListString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<List<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultSortedSetString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<SortedSet<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultCollectionString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<Collection<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultImmutableArray_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableArray<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], [.. row1]);
    }
    
    [Fact]
    public void ReadRow_DefaultImmutableList_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableList<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultImmutableStack_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableStack<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultImmutableQueue_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableQueue<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultImmutableSortedSet_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableSortedSet<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultImmutableHashSet_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ImmutableHashSet<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultFrozenSet_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<FrozenSet<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultObservableCollectionString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<ObservableCollection<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }
    
    [Fact]
    public void ReadRow_DefaultBlockingCollectionString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<BlockingCollection<string>>();
        Assert.Equal(["1", "2", "3", "a", "b", "1", "2", "True", "False", "a", "b", "1", "2", "1,2,3" ], row1);
    }

    [Fact]
    public void ReadRow_MultiMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultiMap.xlsx");
        importer.Configuration.RegisterClassMap(new MultiMapRowMap());

        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<MultiMapRow>();
        Assert.Equal([1, 2, 3], row1.MultiMapName);
        Assert.Equal(new string[] { "a", "b" }, row1.MultiMapIndex);
        Assert.Equal(new string[] { "a", "b" }, row1.IEnumerableNonGeneric);
        Assert.Equal([1, 2], row1.IEnumerableInt);
        Assert.Equal(new string[] { "a", "b" }, row1.ICollectionNonGeneric);
        Assert.Equal(new string[] { "a", "b" }, row1.ICollectionString);
        Assert.Equal([true, false], row1.ICollectionBool);
        Assert.Equal(new string[] { "a", "b" }, row1.IReadOnlyCollectionString);
        Assert.Equal(new string[] { "a", "b" }, row1.IListNonGeneric);
        Assert.Equal(["a", "b"], row1.IListString);
        Assert.Equal(["a", "b"], row1.IListObject);
        Assert.Equal(["a", "b"], row1.IReadOnlyListString);
        Assert.Equal(new ArrayList { "1", "2" }, row1.ArrayList);
        Assert.Equal(new Queue(new string[] { "1", "2" }), row1.QueueNonGeneric);
        Assert.Equal(new Stack(new string[] { "1", "2" }), row1.StackNonGeneric);
        Assert.Equal(new string[] { "1", "2" }, row1.ListString);
        Assert.Equal(new string[] { "1", "2" }, row1._concreteICollection);
        Assert.Equal(new string[] { "1", "2" }, row1.CollectionString);
        Assert.Equal(new string[] { "1", "2" }, row1.ImmutableArrayString);
        Assert.Equal(new string[] { "1", "2" }, row1.ImmutableListString);
        Assert.Equal(new string[] { "1", "2" }, row1.ImmutableStackString);
        Assert.Equal(new string[] { "1", "2" }, row1.ImmutableQueueString);
        Assert.Equal(new string[] { "1", "2" }, row1.ImmutableSortedSetString);
        Assert.Equal(new string[] { "1", "2" }, row1.ImmutableHashSetString);
        Assert.Equal(new string[] { "1", "2" }, row1.FrozenSetString);
        Assert.Equal(new string[] { "1", "2" }, row1.ObservableCollectionString);
        Assert.Equal(new string[] { "1", "2" }, row1.CustomObservableCollectionString);
        Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.a, ObservableCollectionEnum.b }, row1.CustomObservableCollectionEnum);

        var row2 = sheet.ReadRow<MultiMapRow>();
        Assert.Equal([1, -1, 3], row2.MultiMapName);
        Assert.Equal(new string?[] { null, null }, row2.MultiMapIndex);
        Assert.Equal(new string[] { "c", "d" }, row2.IEnumerableNonGeneric);
        Assert.Equal([0, 0], row2.IEnumerableInt);
        Assert.Equal(new string[] { "c", "d" }, row2.ICollectionNonGeneric);
        Assert.Equal(new string[] { "c", "d" }, row2.ICollectionString);
        Assert.Equal([false, true], row2.ICollectionBool);
        Assert.Equal(new string[] { "c", "d" }, row2.IReadOnlyCollectionString);
        Assert.Equal(new string[] { "c", "d" }, row2.IListNonGeneric);
        Assert.Equal(["c", "d"], row2.IListString);
        Assert.Equal(["c", "d"], row2.IListObject);
        Assert.Equal(["c", "d"], row2.IReadOnlyListString);
        Assert.Equal(new ArrayList { "3", "4" }, row2.ArrayList);
        Assert.Equal(new Queue(new string[] { "3", "4" }), row2.QueueNonGeneric);
        Assert.Equal(new Stack(new string[] { "3", "4" }), row2.StackNonGeneric);
        Assert.Equal(new string[] { "3", "4" }, row2.ListString);
        Assert.Equal(new string[] { "3", "4" }, row2._concreteICollection);
        Assert.Equal(new string[] { "3", "4" }, row2.CollectionString);
        Assert.Equal(new string[] { "3", "4" }, row2.ImmutableArrayString);
        Assert.Equal(new string[] { "3", "4" }, row2.ImmutableListString);
        Assert.Equal(new string[] { "3", "4" }, row2.ImmutableStackString);
        Assert.Equal(new string[] { "3", "4" }, row2.ImmutableQueueString);
        Assert.Equal(new string[] { "3", "4" }, row2.ImmutableSortedSetString);
        Assert.Equal(new string[] { "3", "4" }, row2.ImmutableHashSetString);
        Assert.Equal(new string[] { "3", "4" }, row2.FrozenSetString);
        Assert.Equal(new string[] { "3", "4" }, row2.ObservableCollectionString);
        Assert.Equal(new string[] { "3", "4" }, row2.CustomObservableCollectionString);
        Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.custom, ObservableCollectionEnum.custom }, row2.CustomObservableCollectionEnum);

        var row3 = sheet.ReadRow<MultiMapRow>();
        Assert.Equal([-1, -1, -1], row3.MultiMapName);
        Assert.Equal(new string?[] { null, "d" }, row3.MultiMapIndex);
        Assert.Equal(new string[] { "e", "f" }, row3.IEnumerableNonGeneric);
        Assert.Equal([5, 6], row3.IEnumerableInt);
        Assert.Equal(new string[] { "e", "f" }, row3.ICollectionNonGeneric);
        Assert.Equal(new string[] { "e", "f" }, row3.ICollectionString);
        Assert.Equal([false, false], row3.ICollectionBool);
        Assert.Equal(new string[] { "e", "f" }, row3.IReadOnlyCollectionString);
        Assert.Equal(new string[] { "e", "f" }, row3.IListNonGeneric);
        Assert.Equal(["e", "f"], row3.IListString);
        Assert.Equal(["e", "f"], row3.IListObject);
        Assert.Equal(["e", "f"], row3.IReadOnlyListString);
        Assert.Equal(new ArrayList { "5", "6" }, row3.ArrayList);
        Assert.Equal(new Queue(new string[] { "5", "6" }), row3.QueueNonGeneric);
        Assert.Equal(new Stack(new string[] { "5", "6" }), row3.StackNonGeneric);
        Assert.Equal(new string[] { "5", "6" }, row3.ListString);
        Assert.Equal(new string[] { "5", "6" }, row3._concreteICollection);
        Assert.Equal(new string[] { "5", "6" }, row3.CollectionString);
        Assert.Equal(new string[] { "5", "6" }, row3.ImmutableArrayString);
        Assert.Equal(new string[] { "5", "6" }, row3.ImmutableListString);
        Assert.Equal(new string[] { "5", "6" }, row3.ImmutableStackString);
        Assert.Equal(new string[] { "5", "6" }, row3.ImmutableQueueString);
        Assert.Equal(new string[] { "5", "6" }, row3.ImmutableSortedSetString);
        Assert.Equal(new string[] { "5", "6" }, row3.ImmutableHashSetString);
        Assert.Equal(new string[] { "5", "6" }, row3.FrozenSetString);
        Assert.Equal(new string[] { "5", "6" }, row3.ObservableCollectionString);
        Assert.Equal(new string[] { "5", "6" }, row3.CustomObservableCollectionString);
        Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.custom, ObservableCollectionEnum.custom }, row3.CustomObservableCollectionEnum);

        var row4 = sheet.ReadRow<MultiMapRow>();
        Assert.Equal([-2, -2, 3], row4.MultiMapName);
        Assert.Equal(new string?[] { "d", null }, row4.MultiMapIndex);
        Assert.Equal(new string[] { "g", "h" }, row4.IEnumerableNonGeneric);
        Assert.Equal([7, 8], row4.IEnumerableInt);
        Assert.Equal(new string[] { "g", "h" }, row4.ICollectionNonGeneric);
        Assert.Equal(new string[] { "g", "h" }, row4.ICollectionString);
        Assert.Equal([false, true], row4.ICollectionBool);
        Assert.Equal(new string[] { "g", "h" }, row4.IReadOnlyCollectionString);
        Assert.Equal(new string[] { "g", "h" }, row4.IListNonGeneric);
        Assert.Equal(["g", "h"], row4.IListString);
        Assert.Equal(["g", "h"], row4.IListObject);
        Assert.Equal(["g", "h"], row4.IReadOnlyListString);
        Assert.Equal(new ArrayList { "7", "8" }, row4.ArrayList);
        Assert.Equal(new Queue(new string[] { "7", "8" }), row4.QueueNonGeneric);
        Assert.Equal(new Stack(new string[] { "7", "8" }), row4.StackNonGeneric);
        Assert.Equal(new string[] { "7", "8" }, row4.ListString);
        Assert.Equal(new string[] { "7", "8" }, row4._concreteICollection);
        Assert.Equal(new string[] { "7", "8" }, row4.CollectionString);
        Assert.Equal(new string[] { "7", "8" }, row4.ImmutableArrayString);
        Assert.Equal(new string[] { "7", "8" }, row4.ImmutableListString);
        Assert.Equal(new string[] { "7", "8" }, row4.ImmutableStackString);
        Assert.Equal(new string[] { "7", "8" }, row4.ImmutableQueueString);
        Assert.Equal(new string[] { "7", "8" }, row4.ImmutableSortedSetString);
        Assert.Equal(new string[] { "7", "8" }, row4.ImmutableHashSetString);
        Assert.Equal(new string[] { "7", "8" }, row4.FrozenSetString);
        Assert.Equal(new string[] { "7", "8" }, row4.ObservableCollectionString);
        Assert.Equal(new string[] { "7", "8" }, row4.CustomObservableCollectionString);
        Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.custom, ObservableCollectionEnum.custom }, row4.CustomObservableCollectionEnum);
    }

    private class MultiMapRow
    {
        public int[] MultiMapName { get; set; } = default!;
        public CustomList MultiMapIndex { get; set; } = default!;
        public IEnumerable<int> IEnumerableInt { get; set; } = default!;
        public IEnumerable IEnumerableNonGeneric { get; set; } = default!;
        public ICollection ICollectionNonGeneric { get; set; } = default!;
        public ICollection<string> ICollectionString { get; set; } = default!;
        public ICollection<bool> ICollectionBool { get; set; } = default!;
        public IReadOnlyCollection<string> IReadOnlyCollectionString { get; set; } = default!;
        public IList IListNonGeneric { get; set; } = default!;
        public IList<string> IListString { get; set; } = default!;
        public IList<object> IListObject { get; set; } = default!;
        public IReadOnlyList<string> IReadOnlyListString { get; set; } = default!;
        public ArrayList ArrayList { get; set; } = default!;
        public Queue QueueNonGeneric { get; set; } = default!;
        public Stack StackNonGeneric { get; set; } = default!;
        public List<string> ListString { get; set; } = default!;
        public SortedSet<string> _concreteICollection = default!;
        public Collection<string> CollectionString { get; set; } = default!;
        public ImmutableArray<string> ImmutableArrayString { get; set; } = default!;
        public ImmutableList<string> ImmutableListString { get; set; } = default!;
        public ImmutableList<string> ImmutableStackString { get; set; } = default!;
        public ImmutableList<string> ImmutableQueueString { get; set; } = default!;
        public ImmutableList<string> ImmutableSortedSetString { get; set; } = default!;
        public ImmutableList<string> ImmutableHashSetString { get; set; } = default!;
        public FrozenSet<string> FrozenSetString { get; set; } = default!;
        public ObservableCollection<string> ObservableCollectionString { get; set; } = default!;
        public CustomObservableCollection CustomObservableCollectionString { get; set; } = default!;
        public CustomEnumObservableCollection CustomObservableCollectionEnum { get; set; } = default!;
    }

    private class MultiMapRowMap : ExcelClassMap<MultiMapRow>
    {
        public MultiMapRowMap()
        {
            Map(p => p.MultiMapName)
                .WithColumnNames("MultiMapName1", "MultiMapName2", "MultiMapName3")
                .WithElementMap(e => e
                    .WithEmptyFallback(-1)
                    .WithInvalidFallback(-2)
                );

            Map<string>(p => p.MultiMapIndex)
                .WithColumnIndices(3, 4);

            MapList<string>(p => p.IEnumerableNonGeneric)
                .WithColumnNames("IListString1", "IListString2");

            Map(p => p.IEnumerableInt)
                .WithColumnNames(new List<string> { "IEnumerableInt1", "IEnumerableInt2" })
                .WithElementMap(e => e
                    .WithValueFallback(default(int))
                );

            MapList<string>(p => p.ICollectionNonGeneric)
                .WithColumnNames("IListString1", "IListString2");

            MapList<string>(p => p.ICollectionString)
                .WithColumnNames("IListString1", "IListString2");

            Map(p => p.ICollectionBool)
                .WithColumnIndices(new List<int> { 7, 8 })
                .WithElementMap(e => e
                    .WithValueFallback(default(bool))
                );

            MapList<string>(p => p.IReadOnlyCollectionString)
                .WithColumnNames("IListString1", "IListString2");

            MapList<string>(p => p.IListNonGeneric)
                .WithColumnNames("IListString1", "IListString2");

            Map(p => p.IListString)
                .WithColumnNames("IListString1", "IListString2");

            Map(p => p.IListObject)
                .WithColumnNames("IListString1", "IListString2");

            Map(p => p.IReadOnlyListString)
                .WithColumnNames("IListString1", "IListString2");

            MapList<string>(p => p.ArrayList)
                .WithColumnNames("ListString1", "ListString2");

            MapList<string>(p => p.QueueNonGeneric)
                .WithColumnNames("ListString1", "ListString2");

            MapList<string>(p => p.StackNonGeneric)
                .WithColumnNames("ListString1", "ListString2");

            Map(p => p.ListString)
                .WithColumnNames("ListString1", "ListString2");

            Map(p => (ICollection<string>)p._concreteICollection)
                .WithColumnNames("ListString1", "ListString2");

            Map(p => p.CollectionString)
                .WithColumnNames("ListString1", "ListString2");
                
            MapList<string>(p => p.ImmutableArrayString)
                .WithColumnNames("ListString1", "ListString2");
                
            MapList<string>(p => p.ImmutableListString)
                .WithColumnNames("ListString1", "ListString2");
                
            MapList<string>(p => p.ImmutableStackString)
                .WithColumnNames("ListString1", "ListString2");
                
            MapList<string>(p => p.ImmutableQueueString)
                .WithColumnNames("ListString1", "ListString2");
                
            MapList<string>(p => p.ImmutableSortedSetString)
                .WithColumnNames("ListString1", "ListString2");
                
            MapList<string>(p => p.ImmutableHashSetString)
                .WithColumnNames("ListString1", "ListString2");
                
            MapList<string>(p => p.FrozenSetString)
                .WithColumnNames("ListString1", "ListString2");

            Map(p => p.ObservableCollectionString)
                .WithColumnNames("ListString1", "ListString2");

            Map<string>(p => p.CustomObservableCollectionString)
                .WithColumnNames("ListString1", "ListString2");

            Map<ObservableCollectionEnum>(p => p.CustomObservableCollectionEnum)
                .WithColumnNames("MultiMapIndex1", "MultiMapIndex2")
                .WithElementMap(e => e
                    .WithValueFallback(ObservableCollectionEnum.custom)
                );
        }
    }

    public interface INonGenericInteface { }
    public interface IGenericInterface<T> { }
    public interface IMultipleGenericInterface<T, U>{ }

    public class CustomList : INonGenericInteface, IGenericInterface<CustomList>, IList<string>, IMultipleGenericInterface<string, int>
    {
        private IList<string> Inner { get; } = new List<string>();

        public string this[int index]
        {
            get => Inner[0];
            set => Inner[0] = value;
        }

        public int Count => Inner.Count;

        public bool IsReadOnly => Inner.IsReadOnly;

        public void Add(string item) => Inner.Add(item);

        public void Clear() => Inner.Clear();

        public bool Contains(string item) => Inner.Contains(item);

        public void CopyTo(string[] array, int arrayIndex) => Inner.CopyTo(array, arrayIndex);

        public IEnumerator<string> GetEnumerator() => Inner.GetEnumerator();

        public int IndexOf(string item) => Inner.IndexOf(item);

        public void Insert(int index, string item) => Inner.Insert(index, item);

        public bool Remove(string item) => Inner.Remove(item);

        public void RemoveAt(int index) => Inner.RemoveAt(index);

        IEnumerator IEnumerable.GetEnumerator() => Inner.GetEnumerator();
    }

    private class CustomObservableCollection : ObservableCollection<string>
    {
    }

    private class CustomEnumObservableCollection : ObservableCollection<ObservableCollectionEnum>
    {
    }

    private enum ObservableCollectionEnum
    {
        a,
        b,
        custom
    }
}
