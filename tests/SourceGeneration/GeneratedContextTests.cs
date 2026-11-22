using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// A context written by the source generator maps rows as <c>ReadRows&lt;T&gt;()</c> does. Rows
/// here are read both ways and compared.
/// </summary>
public class GeneratedContextTests
{
    [Fact]
    public void Default_Get_ReturnsSameInstance()
    {
        Assert.Same(GeneratedContext.Default, GeneratedContext.Default);
    }

    [Fact]
    public void ReadRows_Values_MatchesAutoMapping()
    {
        const string Csv = "Name,Age,Joined,Score,Active,Id,Day,MaybeDay\nAl,3,2024-03-04,1.5,true,00000000-0000-0000-0000-000000000001,Monday,\nBo,,,2,false,00000000-0000-0000-0000-000000000002,Friday,Sunday\n";
        AssertSame(Csv, GeneratedContext.Default.ValueRow, r => $"{r.Name}/{r.Age}/{r.Joined:yyyy-MM-dd}/{r.Score}/{r.Active}/{r.Id}/{r.Day}/{r.MaybeDay}");
    }

    [Fact]
    public void ReadRows_Collections_MatchesAutoMapping()
    {
        const string Csv = "Name,A,B,Tags,Days\nx,1,2,a;b,Monday;Friday\n";
        AssertSame(Csv, GeneratedContext.Default.CollectionRow, r => $"{string.Join(",", r.Values)}/{string.Join(",", r.Tags)}/{string.Join(",", r.Days)}/{string.Join(",", r.ReadOnly)}");
    }

    [Fact]
    public void ReadRows_NestedObjects_MatchesAutoMapping()
    {
        const string Csv = "Name,Street,City,X,Y\nAl,High St,Leeds,1,2\n";
        AssertSame(Csv, GeneratedContext.Default.NestedRow, r => $"{r.Name}/{r.Address.Street}/{r.Address.City}/{r.Point.X}/{r.Point.Y}");
    }

    [Fact]
    public void ReadRows_Record_MatchesAutoMapping()
    {
        const string Csv = "Sku,Price,Size\nA1,2.5,Large\n";
        AssertSame(Csv, GeneratedContext.Default.RecordRow, r => r.ToString());
    }

    [Fact]
    public void ReadRows_DictionaryRow_MatchesAutoMapping()
    {
        const string Csv = "A,B\n1,2\n";
        AssertSame(Csv, GeneratedContext.Default.DictionaryStringInt32, r => string.Join(",", r.Select(kv => $"{kv.Key}={kv.Value}")));
    }

    [Fact]
    public void ReadRows_ArrayRow_MatchesAutoMapping()
    {
        const string Csv = "A,B\n1,2\n";
        AssertSame(Csv, GeneratedContext.Default.Int32Array, r => string.Join(",", r));
    }

#if LIBRARY_NET8
    [Fact]
    public void ReadRows_Parsable_MatchesAutoMapping()
    {
        const string Csv = "Code\nab\n";
        AssertSame(Csv, GeneratedContext.Default.ParsableRow, r => r.Code.Value);
    }
#endif

    [Fact]
    public void ReadRows_IgnoredMember_MatchesAutoMapping()
    {
        const string Csv = "Name\nAl\n";
        AssertSame(Csv, GeneratedContext.Default.IgnoredRow, r => $"{r.Name}/{r.Ignored is null}");
    }

    [Fact]
    public void GetClassMap_UserProperty_KeepsUserProperty()
    {
        // A member with the property's name isn't replaced: the class map is still given by GetClassMap.
        Assert.Equal("mine", PropertyContext.Default.Named);
        using var importer = Helpers.GetCsvImporter("Name\nAl\n");
        Assert.Equal("Al", importer.ReadSheet().ReadRows(PropertyContext.Default.GetClassMap<Named>()).Single().Name);
    }

    [Fact]
    public void Nested_Context_Generated()
    {
        using var importer = Helpers.GetCsvImporter("Name\nAl\n");
        Assert.Equal("Al", importer.ReadSheet().ReadRows(Outer.NestedContext.Default.Named).Single().Name);
    }

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Compares with auto-mapping.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Compares with auto-mapping.")]
    private static void AssertSame<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(string csv, ExcelClassMap<T> classMap, Func<T, string> describe)
    {
        using var expectedImporter = Helpers.GetCsvImporter(csv);
        var expected = expectedImporter.ReadSheet().ReadRows<T>().Select(describe).ToArray();

        using var importer = Helpers.GetCsvImporter(csv);
        var actual = importer.ReadSheet().ReadRows(classMap).Select(describe).ToArray();
        Assert.NotEmpty(actual);
        Assert.Equal(expected, actual);
    }
}

[ExcelMappable(typeof(ValueRow))]
[ExcelMappable(typeof(CollectionRow))]
[ExcelMappable(typeof(NestedRow))]
[ExcelMappable(typeof(RecordRow))]
[ExcelMappable(typeof(Dictionary<string, int>))]
[ExcelMappable(typeof(int[]))]
#if LIBRARY_NET8
[ExcelMappable(typeof(ParsableRow))]
#endif
[ExcelMappable(typeof(IgnoredRow))]
public partial class GeneratedContext : ExcelMappingContext
{
}

[ExcelMappable(typeof(Named))]
public partial class PropertyContext : ExcelMappingContext
{
    public string Named => "mine";
}

public partial class Outer
{
    [ExcelMappable(typeof(Named))]
    public partial class NestedContext : ExcelMappingContext
    {
    }
}

public class Named
{
    public string? Name { get; set; }
}

public class ValueRow
{
    public string? Name { get; set; }
    public int? Age { get; set; }
    [ExcelFormats("yyyy-MM-dd")]
    public DateTime? Joined { get; set; }
    public double Score { get; set; }
    public bool Active { get; set; }
    public Guid Id { get; set; }
    public DayOfWeek Day { get; set; }
    public DayOfWeek? MaybeDay { get; set; }
}

public class CollectionRow
{
    [ExcelColumnNames("A", "B")]
    public List<int> Values { get; set; } = [];

    [ExcelSeparators(';')]
    public string[] Tags { get; set; } = [];

    [ExcelSeparators(';')]
    public DayOfWeek[] Days { get; set; } = [];

    [ExcelColumnNames("A", "B")]
    public ReadOnlyCollection<long> ReadOnly { get; set; } = new([]);
}

public class NestedRow
{
    public string? Name { get; set; }
    public GeneratedAddress Address { get; set; } = new();
    public GeneratedPoint Point { get; set; }
}

public class GeneratedAddress
{
    public string? Street { get; set; }
    public string? City { get; set; }
}

public struct GeneratedPoint
{
    public int X { get; set; }
    public int Y { get; set; }
}

public enum GeneratedSize
{
    [Description("Small")]
    S,
    [Description("Large")]
    L,
}

public record RecordRow(string Sku, decimal Price, GeneratedSize Size);

#if LIBRARY_NET8
public class ParsableRow
{
    public GeneratedCode Code { get; set; }
}

public readonly struct GeneratedCode(string value) : IParsable<GeneratedCode>
{
    public string Value { get; } = value;

    public static GeneratedCode Parse(string s, IFormatProvider? provider) => new(s.ToUpperInvariant());

    public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, out GeneratedCode result)
    {
        result = new GeneratedCode(s?.ToUpperInvariant() ?? "");
        return s is not null;
    }
}
#endif

public class IgnoredRow
{
    public string? Name { get; set; }

    [ExcelIgnore]
    public object? Ignored { get; set; }
}
