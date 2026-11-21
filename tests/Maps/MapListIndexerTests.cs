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
}
