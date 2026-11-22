using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelMapper;

// Maps rows in a native AOT app. Exits with the number of failures.
// - Rows read with a class map from an ExcelMappingContext (ContextCases) are auto-mapped from the
//   types the context registers, and must not warn: the project makes trimming and AOT warnings errors.
// - Class maps registered with RegisterClassMap whose members are built-in types (ClassMapCases)
//   map too, but the reflection-based APIs warn, so those cases suppress the warnings.
var failures = 0;
ContextCases.Run(Check);
ClassMapCases.Run(Check);
Console.WriteLine(failures == 0 ? "All passed." : $"{failures} failed.");
return failures;

void Check(string name, Func<string> run, string expected)
{
    string actual;
    try
    {
        actual = run();
    }
    catch (Exception exception)
    {
        actual = $"{exception.GetType().Name}: {exception.Message}";
    }

    var passed = actual == expected;
    if (!passed)
    {
        failures++;
    }

    Console.WriteLine($"{(passed ? "PASS" : "FAIL")} {name}: {actual}{(passed ? "" : $" (expected {expected})")}");
}

internal static class Csv
{
    public static MemoryStream Of(string content) => new(Encoding.UTF8.GetBytes(content));

    public static ExcelSheet Sheet(string content) => new ExcelImporter(Of(content), ExcelImporterFileType.Csv).ReadSheet();
}

// Rows read with class maps from a context, which works with native AOT and doesn't warn.
internal static class ContextCases
{
    public static void Run(Action<string, Func<string>, string> check)
    {
        var context = AppContext.Default;

        check("Context: auto-mapped values", () => string.Join(";", Csv.Sheet("Name,Age,Joined,Score\nAlice,34,2024-03-04,1.5\nBob,,,2\n").ReadRows(context.Person)
            .Select(p => $"{p.Name}/{p.Age}/{p.Joined:yyyy-MM-dd}/{p.Score}")), "Alice/34/2024-03-04/1.5;Bob///2");

        check("Context: enums", () =>
        {
            var row = Csv.Sheet("Day,Maybe,Size\nMonday,,Large\n").ReadRows(context.Scheduled).Single();
            return $"{row.Day}/{row.Maybe?.ToString() ?? "null"}/{row.Size}";
        }, "Monday/null/L");

        check("Context: parsable", () =>
        {
            var row = Csv.Sheet("Code,Name\nab,CD\n").ReadRows(context.Parsed).Single();
            return $"{row.Code.Value}/{row.Name?.Value}";
        }, "AB/cd");

        check("Context: collections", () =>
        {
            var row = Csv.Sheet("Name,A,B,Tags\nx,1,2,a;b\n").ReadRows(context.Collected).Single();
            return $"{string.Join(",", row.Values)}/{string.Join(",", row.Tags)}/{string.Join(",", row.Days)}";
        }, "1,2/a,b/");

        check("Context: dictionary rows", () => string.Join(",", Csv.Sheet("A,B\n1,2\n").ReadRows(context.DictionaryStringInt32).Single()
            .Select(kv => $"{kv.Key}={kv.Value}")), "A=1,B=2");

        check("Context: nested struct", () =>
        {
            var row = Csv.Sheet("Name,X,Y\np,1,2\n").ReadRows(context.Placed).Single();
            return $"{row.Name}/{row.Point.X}/{row.Point.Y}";
        }, "p/1/2");

        check("Context: record and converter", () =>
        {
            var row = Csv.Sheet("Sku,Amount,Maybe\nA1,2k,3k\n").ReadRows(context.Priced).Single();
            return $"{row.Sku}/{row.Amount}/{row.Maybe}";
        }, "A1/2000/3000");

        check("Context: row results", () =>
        {
            var result = Csv.Sheet("Name,Age,Joined,Score\nAlice,x,,1\n").ReadRowResults(context.Person).Single();
            return string.Join(",", result.Errors.Select(e => e.CellAddress));
        }, "B2");

        check("Context: heading", () =>
        {
            var sheet = Csv.Sheet("Title,,,\nName,Age,Joined,Score\nAl,1,,1\n");
            sheet.FindHeading(context.Person);
            sheet.ValidateHeading(context.Person);
            return sheet.ReadRow(context.Person).Name!;
        }, "Al");

        check("Context: region", () => Csv.Sheet("Title,,,,\n,Name,Age,Joined,Score\n,Al,3,,1\n").Region("B2:E3").ReadRows(context.Person).Single().Age.ToString()!, "3");

        check("Context: unregistered type", () =>
        {
            try
            {
                context.GetClassMap<Unregistered>();
                return "no exception";
            }
            catch (InvalidOperationException)
            {
                return "InvalidOperationException";
            }
        }, "InvalidOperationException");
    }
}

// A context the source generator writes: it registers each row type and every type their members
// use, and adds a property for each row type's class map and a Default instance.
[ExcelMappable(typeof(Person))]
[ExcelMappable(typeof(Scheduled))]
[ExcelMappable(typeof(Parsed))]
[ExcelMappable(typeof(Collected))]
[ExcelMappable(typeof(Dictionary<string, int>))]
[ExcelMappable(typeof(Placed))]
[ExcelMappable(typeof(Priced))]
internal sealed partial class AppContext : ExcelMappingContext
{
}

// Class maps of built-in types, registered with the reflection-based APIs, which warn.
internal static class ClassMapCases
{
    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Checks class maps of built-in types, whose members are kept.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Checks class maps of built-in types, which native AOT has the code for.")]
    public static void Run(Action<string, Func<string>, string> check)
    {
        check("Class map: values", () =>
        {
            using var importer = new ExcelImporter(Csv.Of("I,N,L,D,M,B,T,G,S\n1,,3,1.5,2.25,true,2024-03-04,00000000-0000-0000-0000-000000000001,x\n"), ExcelImporterFileType.Csv);
            importer.Configuration.RegisterClassMap<Values>(c =>
            {
                c.Map(v => v.I).WithColumnName("I");
                c.Map(v => v.N).WithColumnName("N");
                c.Map(v => v.L).WithColumnName("L");
                c.Map(v => v.D).WithColumnName("D");
                c.Map(v => v.M).WithColumnName("M");
                c.Map(v => v.B).WithColumnName("B");
                c.Map(v => v.T).WithColumnName("T").WithFormats("yyyy-MM-dd");
                c.Map(v => v.G).WithColumnName("G");
                c.Map(v => v.S).WithColumnName("S");
            });
            var v = importer.ReadSheet().ReadRows<Values>().Single();
            return $"{v.I}/{v.N?.ToString() ?? "null"}/{v.L}/{v.D}/{v.M}/{v.B}/{v.T:yyyy-MM-dd}/{v.G.ToString()[^1]}/{v.S}";
        }, "1/null/3/1.5/2.25/True/2024-03-04/1/x");

        check("Class map: list", () =>
        {
            using var importer = new ExcelImporter(Csv.Of("A,B\n1,2\n"), ExcelImporterFileType.Csv);
            importer.Configuration.RegisterClassMap<Listed>(c => c.MapList<int>(l => l.Values).WithColumnNames("A", "B"));
            return string.Join(",", importer.ReadSheet().ReadRows<Listed>().Single().Values);
        }, "1,2");

        check("Class map: record", () =>
        {
            using var importer = new ExcelImporter(Csv.Of("Sku,Price\nA1,2.5\n"), ExcelImporterFileType.Csv);
            importer.Configuration.RegisterClassMap<Item>(c =>
            {
                c.Map(i => i.Sku);
                c.Map(i => i.Price);
            });
            return importer.ReadSheet().ReadRows<Item>().Single().ToString();
        }, "Item { Sku = A1, Price = 2.5 }");

        check("Class map: struct", () =>
        {
            using var importer = new ExcelImporter(Csv.Of("X,Y\n1,2\n"), ExcelImporterFileType.Csv);
            importer.Configuration.RegisterClassMap<Point>(c =>
            {
                c.Map(p => p.X);
                c.Map(p => p.Y);
            });
            var point = importer.ReadSheet().ReadRows<Point>().Single();
            return $"{point.X}/{point.Y}";
        }, "1/2");

        check("Class map: single values", () =>
        {
            using var importer = new ExcelImporter(Csv.Of("V\n1\n2\n"), ExcelImporterFileType.Csv);
            importer.Configuration.RegisterClassMap<int>(c => c.Map(v => v));
            return string.Join(",", importer.ReadSheet().ReadRows<int>());
        }, "1,2");

        check("Class map: errors", () =>
        {
            using var importer = new ExcelImporter(Csv.Of("X,Y\n1,a\n"), ExcelImporterFileType.Csv);
            importer.Configuration.RegisterClassMap<Point>(c =>
            {
                c.Map(p => p.X);
                c.Map(p => p.Y);
            });
            var result = importer.ReadSheet().ReadRowResults<Point>().Single();
            return string.Join(",", result.Errors.Select(e => e.CellAddress));
        }, "B2");

        check("Auto-mapped strings", () => string.Join(";", ExcelImporter.ReadRows<Named>(Csv.Of("First,Last\nAda,Lovelace\n"), ExcelImporterFileType.Csv)
            .Select(n => $"{n.First} {n.Last}")), "Ada Lovelace");
    }
}

public class Person
{
    public string? Name { get; set; }
    public int? Age { get; set; }
    [ExcelFormats("yyyy-MM-dd")]
    public DateTime? Joined { get; set; }
    public double Score { get; set; }
}

public class Scheduled
{
    public DayOfWeek Day { get; set; }
    public DayOfWeek? Maybe { get; set; }
    public Size Size { get; set; }
}

public enum Size
{
    [Description("Small")]
    S,
    [Description("Large")]
    L,
}

public class Parsed
{
    public Code Code { get; set; }

    // Parsed without a parameterless constructor.
    public Name? Name { get; set; }
}

public sealed record Name(string Value) : IParsable<Name>
{
    public static Name Parse(string s, IFormatProvider? provider) => new(s.ToLowerInvariant());

    public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, [MaybeNullWhen(false)] out Name result)
    {
        result = s is null ? null : Parse(s, provider);
        return result is not null;
    }
}

public readonly struct Code(string value) : IParsable<Code>
{
    public string Value { get; } = value;

    public static Code Parse(string s, IFormatProvider? provider) => new(s.ToUpperInvariant());

    public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, out Code result)
    {
        result = s is null ? default : new Code(s.ToUpperInvariant());
        return s is not null;
    }
}

public class Collected
{
    [ExcelColumnNames("A", "B")]
    public List<int> Values { get; set; } = [];

    [ExcelSeparators(';')]
    public string[] Tags { get; set; } = [];

    [ExcelOptional]
    public DayOfWeek[] Days { get; set; } = [];
}

public class Placed
{
    public string? Name { get; set; }
    public Point Point { get; set; }
}

public struct Point
{
    public int X { get; set; }
    public int Y { get; set; }
}

public record Priced(string Sku, [property: ExcelConverter(typeof(ThousandsConverter))] decimal Amount, [property: ExcelConverter(typeof(ThousandsConverter))] decimal? Maybe);

public class ThousandsConverter : ICellConverter<decimal>
{
    public CellConversionResult<decimal> Convert(CellValue value, CellContext context)
        => value.Text is { } text && text.EndsWith('k') && decimal.TryParse(text[..^1], out var thousands)
            ? CellConversionResult<decimal>.Success(thousands * 1000)
            : CellConversionResult<decimal>.Invalid();
}

public class Unregistered
{
    public string? Name { get; set; }
}

public class Values
{
    public int I { get; set; }
    public int? N { get; set; }
    public long L { get; set; }
    public double D { get; set; }
    public decimal M { get; set; }
    public bool B { get; set; }
    public DateTime T { get; set; }
    public Guid G { get; set; }
    public string? S { get; set; }
}

public class Listed
{
    public List<int> Values { get; set; } = [];
}

public record Item(string Sku, double Price);

public class Named
{
    public string? First { get; set; }
    public string? Last { get; set; }
}
