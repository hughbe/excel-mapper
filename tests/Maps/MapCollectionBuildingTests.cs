using System.Collections.Concurrent;
using System.Collections.Frozen;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelMapper.Tests;

/// <summary>
/// Collections are built by a builder shared by every sheet and thread, which can reuse a buffer
/// between rows when the result is a copy of it. These check that rows never share storage.
/// </summary>
public class MapCollectionBuildingTests
{
    private const string Csv = "A,B,C\n1,2,3\n4,5,6\n7,8,9\n";

    [Fact]
    public void ReadRows_CollectionsFromReusedBuffers_AreIndependent()
    {
        using var importer = Helpers.GetCsvImporter(Csv);
        var rows = importer.ReadSheet().ReadRows<Collections>().ToArray();

        Assert.Equal([[1, 2, 3], [4, 5, 6], [7, 8, 9]], rows.Select(r => r.Frozen.OrderBy(v => v).ToArray()));
        Assert.Equal([[1, 2, 3], [4, 5, 6], [7, 8, 9]], rows.Select(r => r.Queue.ToArray()));
        Assert.Equal([[3, 2, 1], [6, 5, 4], [9, 8, 7]], rows.Select(r => r.Stack.ToArray()));
        Assert.Equal([[1, 2, 3], [4, 5, 6], [7, 8, 9]], rows.Select(r => r.ImmutableQueue.ToArray()));
        Assert.Equal([[3, 2, 1], [6, 5, 4], [9, 8, 7]], rows.Select(r => r.ImmutableStack.ToArray()));
        Assert.Equal([[1, 2, 3], [4, 5, 6], [7, 8, 9]], rows.Select(r => r.ReadOnly.ToArray()));
        Assert.Equal([[1, 2, 3], [4, 5, 6], [7, 8, 9]], rows.Select(r => r.ImmutableArray.ToArray()));
        Assert.Equal([[1, 2, 3], [4, 5, 6], [7, 8, 9]], rows.Select(r => r.Custom.Values.ToArray()));
    }

    public class Collections
    {
        [ExcelColumnNames("A", "B", "C")]
        public FrozenSet<int> Frozen { get; set; } = FrozenSet<int>.Empty;

        [ExcelColumnNames("A", "B", "C")]
        public Queue<int> Queue { get; set; } = new();

        [ExcelColumnNames("A", "B", "C")]
        public Stack<int> Stack { get; set; } = new();

        [ExcelColumnNames("A", "B", "C")]
        public ImmutableQueue<int> ImmutableQueue { get; set; } = ImmutableQueue<int>.Empty;

        [ExcelColumnNames("A", "B", "C")]
        public ImmutableStack<int> ImmutableStack { get; set; } = ImmutableStack<int>.Empty;

        // Wraps the list it's given, so each row needs its own list.
        [ExcelColumnNames("A", "B", "C")]
        public ReadOnlyCollection<int> ReadOnly { get; set; } = new([]);

        [ExcelColumnNames("A", "B", "C")]
        public ImmutableArray<int> ImmutableArray { get; set; }

        // Built with a parameterless constructor and Add(int), called through a compiled delegate.
        [ExcelColumnNames("A", "B", "C")]
        public AddOnly Custom { get; set; } = new();
    }

    public class AddOnly : IEnumerable<int>
    {
        public List<int> Values { get; } = [];

        public void Add(int value) => Values.Add(value);

        public IEnumerator<int> GetEnumerator() => Values.GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    [Fact]
    public void ReadRows_DictionariesFromReusedBuffers_AreIndependent()
    {
        using var importer = Helpers.GetCsvImporter(Csv);
        var rows = importer.ReadSheet().ReadRows<Dictionaries>().ToArray();

        Assert.Equal([1, 4, 7], rows.Select(r => r.Frozen["A"]));
        Assert.Equal([3, 6, 9], rows.Select(r => r.Sorted["C"]));
        Assert.Equal([2, 5, 8], rows.Select(r => r.ReadOnly["B"]));
        Assert.Equal([1, 4, 7], rows.Select(r => r.Custom.Values["A"]));
    }

    public class Dictionaries
    {
        [ExcelColumnNames("A", "B", "C")]
        public FrozenDictionary<string, int> Frozen { get; set; } = FrozenDictionary<string, int>.Empty;

        [ExcelColumnNames("A", "B", "C")]
        public SortedDictionary<string, int> Sorted { get; set; } = [];

        [ExcelColumnNames("A", "B", "C")]
        public ReadOnlyDictionary<string, int> ReadOnly { get; set; } = new(new Dictionary<string, int>());

        [ExcelColumnNames("A", "B", "C")]
        public AddOnlyDictionary Custom { get; set; } = new();
    }

    public class AddOnlyDictionary : IEnumerable<KeyValuePair<string, int>>
    {
        public Dictionary<string, int> Values { get; } = [];

        public void Add(string key, int value) => Values.Add(key, value);

        public IEnumerator<KeyValuePair<string, int>> GetEnumerator() => Values.GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    [Fact]
    public void ReadRows_ClassMapSharedAcrossThreads_ReturnsExpected()
    {
        // Builders keep no state, and collections created from a copy of their elements gather
        // them in a list of each thread's own, so a class map shared between importers on
        // different threads builds each row's collections separately.
        var csv = "A,B,C\n" + string.Concat(Enumerable.Range(0, 2000).Select(i => $"{i},{i + 1},{i + 2}\n"));

        // One class map instance: auto-mapped from the members' attributes, with two elements
        // mapped one by one.
        var map = new ExcelClassMap<Shared>().AutoMap();
        map.Map(r => r.Indexed[0]).WithColumnName("C");
        map.Map(r => r.Indexed[1]).WithColumnName("A");

        var results = new ConcurrentBag<bool>();
        Parallel.For(0, 8, _ =>
        {
            using var importer = Helpers.GetCsvImporter(csv);
            importer.Configuration.RegisterClassMap(map);
            var i = 0;
            foreach (var row in importer.ReadSheet().ReadRows<Shared>())
            {
                results.Add(
                    row.Values.SequenceEqual([i, i + 1, i + 2]) &&
                    row.Set.SetEquals([i, i + 1, i + 2]) &&
                    row.Immutable.SequenceEqual([i, i + 1, i + 2]) &&
                    row.Frozen.SetEquals([i, i + 1, i + 2]) &&
                    row.Dictionary.Count == 3 && row.Dictionary["A"] == i && row.Dictionary["C"] == i + 2 &&
                    row.Indexed.SequenceEqual([i + 2, i]));
                i++;
            }
        });

        Assert.Equal(16000, results.Count);
        Assert.All(results, Assert.True);
    }

    public class Shared
    {
        [ExcelColumnNames("A", "B", "C")]
        public List<int> Values { get; set; } = [];

        [ExcelColumnNames("A", "B", "C")]
        public ISet<int> Set { get; set; } = new HashSet<int>();

        [ExcelColumnNames("A", "B", "C")]
        public ImmutableList<int> Immutable { get; set; } = [];

        [ExcelColumnNames("A", "B", "C")]
        public FrozenSet<int> Frozen { get; set; } = FrozenSet<int>.Empty;

        [ExcelColumnNames("A", "B", "C")]
        public ImmutableDictionary<string, int> Dictionary { get; set; } = ImmutableDictionary<string, int>.Empty;

        [ExcelIgnore]
        public int[] Indexed { get; set; } = [];
    }

    [Fact]
    public void ReadRows_CollectionConstructorThrows_ReportsTheConstructorsException()
    {
        using var importer = Helpers.GetCsvImporter(Csv);
        var sheet = importer.ReadSheet();
        Assert.Throws<NotSupportedException>(() => sheet.ReadRows<ThrowingRow>().ToArray());
    }

    public class ThrowingRow
    {
        [ExcelColumnNames("A", "B")]
        public ThrowingCollection Values { get; set; } = new([]);
    }

    public class ThrowingCollection : IEnumerable<int>
    {
        public ThrowingCollection(IList<int> values) => throw new NotSupportedException();

        public IEnumerator<int> GetEnumerator() => throw new NotImplementedException();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    [Fact]
    public void ReadRows_CollectionsFromConstructorsWithDuplicates_KeepEveryValue()
    {
        // Collections built by passing the values to a constructor keep duplicates and order: only a
        // constructor that needs a set is given one.
        using var importer = Helpers.GetCsvImporter("A,B,C\n2,1,2\n3,3,1\n");
        var rows = importer.ReadSheet().ReadRows<ConstructedCollections>().ToArray();
        Assert.Equal([[2, 1, 2], [3, 3, 1]], rows.Select(r => r.Queue.ToArray()));
        Assert.Equal([[2, 1, 2], [1, 3, 3]], rows.Select(r => r.Stack.ToArray()));
        Assert.Equal([[2, 1, 2], [3, 3, 1]], rows.Select(r => r.ConcurrentQueue.ToArray()));
        Assert.Equal([[2, 1, 2], [1, 3, 3]], rows.Select(r => r.ConcurrentStack.ToArray()));
        Assert.Equal([[1, 2, 2], [1, 3, 3]], rows.Select(r => r.ConcurrentBag.OrderBy(v => v).ToArray()));
        Assert.Equal([[2, 1, 2], [3, 3, 1]], rows.Select(r => r.FromCollection.Values));
        Assert.Equal([[2, 1, 2], [3, 3, 1]], rows.Select(r => r.FromReadOnlyList.Values));
        Assert.Equal([[2, 1], [3, 1]], rows.Select(r => r.FromSet.Values));
    }

    public class ConstructedCollections
    {
        [ExcelColumnNames("A", "B", "C")]
        public Queue<int> Queue { get; set; } = new();

        [ExcelColumnNames("A", "B", "C")]
        public Stack<int> Stack { get; set; } = new();

        [ExcelColumnNames("A", "B", "C")]
        public ConcurrentQueue<int> ConcurrentQueue { get; set; } = new();

        [ExcelColumnNames("A", "B", "C")]
        public ConcurrentStack<int> ConcurrentStack { get; set; } = new();

        [ExcelColumnNames("A", "B", "C")]
        public ConcurrentBag<int> ConcurrentBag { get; set; } = new();

        [ExcelColumnNames("A", "B", "C")]
        public FromCollectionType FromCollection { get; set; } = new([]);

        [ExcelColumnNames("A", "B", "C")]
        public FromReadOnlyListType FromReadOnlyList { get; set; } = new([]);

        [ExcelColumnNames("A", "B", "C")]
        public FromSetType FromSet { get; set; } = new(new HashSet<int>());
    }

    // Collections with only a constructor that takes the values.
    public class FromCollectionType(ICollection<int> values) : IEnumerable<int>
    {
        public int[] Values { get; } = [.. values];
        public IEnumerator<int> GetEnumerator() => ((IEnumerable<int>)Values).GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public class FromReadOnlyListType(IReadOnlyList<int> values) : IEnumerable<int>
    {
        public int[] Values { get; } = [.. values];
        public IEnumerator<int> GetEnumerator() => ((IEnumerable<int>)Values).GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public class FromSetType(ISet<int> values) : IEnumerable<int>
    {
        public int[] Values { get; } = [.. values];
        public IEnumerator<int> GetEnumerator() => ((IEnumerable<int>)Values).GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
