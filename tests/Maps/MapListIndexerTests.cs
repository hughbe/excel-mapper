using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// Mapping elements of lists by index (for example <c>Map(r =&gt; r.Values[2])</c>) for each kind of
/// list type: the list is built from the mapped indices, with default values in any gaps.
/// </summary>
public class MapListIndexerTests
{
    [Fact]
    public void ReadRows_InterfaceListIndexers_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("A,B\n1,2\n3,4\n");
        importer.Configuration.RegisterClassMap<InterfaceLists>(c =>
        {
            c.Map(r => r.List[0]).WithColumnName("A");
            c.Map(r => r.List[2]).WithColumnName("B");
            c.Map(r => r.ReadOnlyList[1]).WithColumnName("A");
        });
        var rows = importer.ReadSheet().ReadRows<InterfaceLists>().ToArray();
        Assert.Equal([1, 0, 2], rows[0].List);
        Assert.Equal([3, 0, 4], rows[1].List);
        Assert.Equal([0, 1], rows[0].ReadOnlyList);
        Assert.Equal([0, 3], rows[1].ReadOnlyList);
    }

    [Fact]
    public void ReadRows_ConcreteListIndexers_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("A,B\n1,2\n3,4\n");
        importer.Configuration.RegisterClassMap<ConcreteLists>(c =>
        {
            c.Map(r => r.Observable[1]).WithColumnName("A");
            c.Map(r => r.Observable[0]).WithColumnName("B");
            c.Map(r => r.ReadOnly[1]).WithColumnName("B");
            c.Map(r => r.Immutable[0]).WithColumnName("A");
        });
        var rows = importer.ReadSheet().ReadRows<ConcreteLists>().ToArray();
        Assert.Equal([2, 1], rows[0].Observable);
        Assert.Equal([4, 3], rows[1].Observable);
        Assert.Equal([0, 2], rows[0].ReadOnly);
        Assert.Equal([0, 4], rows[1].ReadOnly);
        Assert.Equal([1], rows[0].Immutable);
        Assert.Equal([3], rows[1].Immutable);
    }

    public class InterfaceLists
    {
        public IList<int> List { get; set; } = [];
        public IReadOnlyList<int> ReadOnlyList { get; set; } = [];
    }

    public class ConcreteLists
    {
        public ReadOnlyObservableCollection<int> Observable { get; set; } = new([]);
        public ReadOnlyCollection<int> ReadOnly { get; set; } = new([]);
        public ImmutableList<int> Immutable { get; set; } = [];
    }

    [Fact]
    public void ReadRows_IndexesFromLoopVariable_ReturnsExpected()
    {
        // An index can be a variable the expression captures, such as a loop's.
        using var importer = Helpers.GetCsvImporter("Q1,Q2,Q3\n1,2,3\n");
        importer.Configuration.RegisterClassMap<CapturedIndexRow>(c =>
        {
            for (var i = 0; i < 3; i++)
            {
                c.Map(r => r.Array[i]).WithColumnName($"Q{i + 1}");
                c.Map(r => r.List[i]).WithColumnName($"Q{3 - i}");
            }
        });

        var row = importer.ReadSheet().ReadRows<CapturedIndexRow>().Single();
        Assert.Equal([1, 2, 3], row.Array);
        Assert.Equal([3, 2, 1], row.List);
    }

    [Fact]
    public void ReadRows_KeysAndGridIndexesFromVariablesAndFields_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Maths,English,A\n90,80,x\n");
        var keys = new[] { "Maths", "English" };
        short row = 0;
        importer.Configuration.RegisterClassMap<CapturedIndexRow>(c =>
        {
            foreach (var key in keys)
            {
                c.Map(r => r.Scores[key]);
            }

            c.Map(r => r.Grid[row, s_column]).WithColumnName("A");
        });

        var result = importer.ReadSheet().ReadRows<CapturedIndexRow>().Single();
        Assert.Equal(new Dictionary<string, int> { ["Maths"] = 90, ["English"] = 80 }, result.Scores);
        Assert.Equal("x", result.Grid[0, 1]);
    }

    private static readonly int s_column = 1;

    [Fact]
    public void Map_NegativeCapturedIndex_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<CapturedIndexRow>();
        var index = -1;
        var exception = Assert.Throws<ArgumentException>("expression", () => map.Map(r => r.Array[index]));
        Assert.StartsWith("Array and list indices must be non-negative. Received: -1 in ", exception.Message);
    }

    [Fact]
    public void Map_IndexFromRow_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<CapturedIndexRow>();
        var exception = Assert.Throws<ArgumentException>("expression", () => map.Map(r => r.Array[r.Array.Length]));
        Assert.StartsWith("An index must be a constant, or a variable or field captured by the expression", exception.Message);
    }

    public class CapturedIndexRow
    {
        public int[] Array { get; set; } = [];

        public List<int> List { get; set; } = [];

        public Dictionary<string, int> Scores { get; set; } = [];

        public string?[,] Grid { get; set; } = new string?[1, 2];
    }
}
