using System.Text;
using ExcelMapper;

// Maps rows in a native AOT app. Class maps registered with RegisterClassMap, whose members are
// built-in types (numbers, strings, dates, booleans, GUIDs, nullable values and lists of them), map
// here: their types are known when the app is compiled, so native AOT has the code they need.
// Auto-mapped members of value types, enums and other types found at run time aren't supported yet.
// Exits with the number of failures.
var failures = 0;

Check("Values", () =>
{
    using var importer = new ExcelImporter(Csv("I,N,L,D,M,B,T,G,S\n1,,3,1.5,2.25,true,2024-03-04,00000000-0000-0000-0000-000000000001,x\n"), ExcelImporterFileType.Csv);
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

Check("List", () =>
{
    using var importer = new ExcelImporter(Csv("A,B\n1,2\n"), ExcelImporterFileType.Csv);
    importer.Configuration.RegisterClassMap<Listed>(c => c.MapList<int>(l => l.Values).WithColumnNames("A", "B"));
    return string.Join(",", importer.ReadSheet().ReadRows<Listed>().Single().Values);
}, "1,2");

Check("Record", () =>
{
    using var importer = new ExcelImporter(Csv("Sku,Price\nA1,2.5\n"), ExcelImporterFileType.Csv);
    importer.Configuration.RegisterClassMap<Item>(c =>
    {
        c.Map(i => i.Sku);
        c.Map(i => i.Price);
    });
    return importer.ReadSheet().ReadRows<Item>().Single().ToString();
}, "Item { Sku = A1, Price = 2.5 }");

Check("Struct", () =>
{
    using var importer = new ExcelImporter(Csv("X,Y\n1,2\n"), ExcelImporterFileType.Csv);
    importer.Configuration.RegisterClassMap<Point>(c =>
    {
        c.Map(p => p.X);
        c.Map(p => p.Y);
    });
    var point = importer.ReadSheet().ReadRows<Point>().Single();
    return $"{point.X}/{point.Y}";
}, "1/2");

Check("Single values", () =>
{
    using var importer = new ExcelImporter(Csv("V\n1\n2\n"), ExcelImporterFileType.Csv);
    importer.Configuration.RegisterClassMap<int>(c => c.Map(v => v));
    return string.Join(",", importer.ReadSheet().ReadRows<int>());
}, "1,2");

Check("Errors", () =>
{
    using var importer = new ExcelImporter(Csv("X,Y\n1,a\n"), ExcelImporterFileType.Csv);
    importer.Configuration.RegisterClassMap<Point>(c =>
    {
        c.Map(p => p.X);
        c.Map(p => p.Y);
    });
    var result = importer.ReadSheet().ReadRowResults<Point>().Single();
    return string.Join(",", result.Errors.Select(e => e.CellAddress));
}, "B2");

Check("Strings auto-mapped", () => string.Join(";", ExcelImporter.ReadRows<Named>(Csv("First,Last\nAda,Lovelace\n"), ExcelImporterFileType.Csv)
    .Select(n => $"{n.First} {n.Last}")), "Ada Lovelace");

Console.WriteLine(failures == 0 ? "All passed." : $"{failures} failed.");
return failures;

static MemoryStream Csv(string content) => new(Encoding.UTF8.GetBytes(content));

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

public struct Point
{
    public int X { get; set; }
    public int Y { get; set; }
}

public class Named
{
    public string? First { get; set; }
    public string? Last { get; set; }
}
