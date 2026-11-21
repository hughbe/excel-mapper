using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelMapper;

// Maps rows in a trimmed app. The mapped types' own members are kept by the trimming annotations on
// ReadRows<T>, RegisterClassMap<T> and ExcelClassMap<T>; nested types need their members kept by the
// app (here with DynamicDependency). Exits with the number of failures.
var failures = 0;

Check("Class", () => string.Join(";", ExcelImporter.ReadRows<Person>(Csv("Name,Age,Joined\nAlice,34,2024-03-04\nBob,,\n"), ExcelImporterFileType.Csv)
    .Select(p => $"{p.Name}/{p.Age}/{p.Joined:yyyy-MM-dd}")), "Alice/34/2024-03-04;Bob//");

Check("Record", () => ExcelImporter.ReadRows<Item>(Csv("Sku,Price\nA1,2.5\n"), ExcelImporterFileType.Csv).Single().ToString(), "Item { Sku = A1, Price = 2.5 }");

Check("Class map", () =>
{
    using var importer = new ExcelImporter(Csv("Ref,Amount\nX,10\n"), ExcelImporterFileType.Csv);
    importer.Configuration.RegisterClassMap<Order>(c =>
    {
        c.Map(o => o.Id).WithColumnName("Ref");
        c.Map(o => o.Total).WithColumnName("Amount");
    });
    var order = importer.ReadSheet().ReadRows<Order>().Single();
    return $"{order.Id}/{order.Total}";
}, "X/10");

Check("Split list", () => string.Join(",", ExcelImporter.ReadRows<Tagged>(Csv("Name,Tags\nx,a;b\n"), ExcelImporterFileType.Csv).Single().Tags), "a,b");

Check("Dictionary", () => string.Join(",", ExcelImporter.ReadRows<Dictionary<string, string>>(Csv("A,B\n1,2\n"), ExcelImporterFileType.Csv)
    .Single().Select(kv => $"{kv.Key}={kv.Value}")), "A=1,B=2");

Check("Enum", () => ExcelImporter.ReadRows<Scheduled>(Csv("Day\nMonday\n"), ExcelImporterFileType.Csv).Single().Day.ToString(), "Monday");

Check("Enum description", () => ExcelImporter.ReadRows<Sized>(Csv("Size\nLarge\n"), ExcelImporterFileType.Csv).Single().Size.ToString(), "L");

Check("Cell converter", () =>
{
    var row = ExcelImporter.ReadRows<Converted>(Csv("Amount,Maybe\n2k,3k\n"), ExcelImporterFileType.Csv).Single();
    return $"{row.Amount}/{row.Maybe}";
}, "2000/3000");

Check("Cell transformer and fallback", () =>
{
    var rows = ExcelImporter.ReadRows<Cleaned>(Csv("Code,Count\nab,x\n"), ExcelImporterFileType.Csv).Single();
    return $"{rows.Code}/{rows.Count}";
}, "AB/-1");

Check("Column selector", () =>
{
    var row = ExcelImporter.ReadRows<Selected>(Csv("Name,Amount (GBP),Q1,Q2\na,5,1,2\n"), ExcelImporterFileType.Csv).Single();
    return $"{row.Amount}/{string.Join(",", row.Quarters)}";
}, "5/1,2");

Check("Nullable numbers", () =>
{
    var row = ExcelImporter.ReadRows<Counted>(Csv("Count,Amount\n,2.5\n"), ExcelImporterFileType.Csv).Single();
    return $"{row.Count?.ToString() ?? "null"}/{row.Amount}";
}, "null/2.5");

Check("Parsable", () => ExcelImporter.ReadRows<Parsed>(Csv("Code\nab\n"), ExcelImporterFileType.Csv).Single().Code.Value, "AB");

Check("Region", () =>
{
    using var importer = new ExcelImporter(Csv("Title,,,\n,Ref,Amount,\n,X,10,\n"), ExcelImporterFileType.Csv);
    var order = importer.ReadSheet().Region("B2:C3").ReadRows<Order>().Single();
    return $"{order.Total}";
}, "10");

Check("Nested object", () =>
{
    KeepNestedTypes();
    var customer = ExcelImporter.ReadRows<Customer>(Csv("Name,Street\nAl,High St\n"), ExcelImporterFileType.Csv).Single();
    return $"{customer.Name}/{customer.Address?.Street}";
}, "Al/High St");

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

[DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicProperties, typeof(Address))]
// Member types that are mapped by their interfaces (here IParsable<T>) need their interfaces kept.
[DynamicDependency(DynamicallyAccessedMemberTypes.Interfaces | DynamicallyAccessedMemberTypes.PublicMethods, typeof(Code))]
static void KeepNestedTypes()
{
}

public class Person
{
    public string? Name { get; set; }
    public int? Age { get; set; }
    [ExcelFormats("yyyy-MM-dd")]
    public DateTime? Joined { get; set; }
}

public record Item(string Sku, double Price);

public class Order
{
    [ExcelColumnNames("Id", "Ref")]
    public string? Id { get; set; }
    [ExcelColumnNames("Total", "Amount")]
    public decimal Total { get; set; }
}

public class Tagged
{
    public string? Name { get; set; }
    [ExcelSeparators(';')]
    public string[] Tags { get; set; } = [];
}

public class Scheduled
{
    public DayOfWeek Day { get; set; }
}

public class Customer
{
    public string? Name { get; set; }
    public Address? Address { get; set; }
}

public class Address
{
    public string? Street { get; set; }
}

public enum Size
{
    [System.ComponentModel.Description("Small")]
    S,
    [System.ComponentModel.Description("Large")]
    L,
}

public class Sized
{
    public Size Size { get; set; }
}

public class Counted
{
    public int? Count { get; set; }
    public decimal Amount { get; set; }
}

public class Parsed
{
    public Code Code { get; set; }
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

public class Converted
{
    [ExcelConverter(typeof(ThousandsConverter))]
    public int Amount { get; set; }

    [ExcelConverter(typeof(ThousandsConverter))]
    public int? Maybe { get; set; }
}

public class ThousandsConverter : ICellConverter<int>
{
    public CellConversionResult<int> Convert(CellValue value, CellContext context)
        => value.Text is { } text && text.EndsWith('k') && int.TryParse(text[..^1], out var thousands)
            ? CellConversionResult<int>.Success(thousands * 1000)
            : CellConversionResult<int>.Invalid();
}

public class Cleaned
{
    [ExcelTransformer(typeof(UpperTransformer))]
    public string? Code { get; set; }

    [ExcelInvalidFallback(typeof(MinusOneFallback))]
    public int? Count { get; set; }
}

public class UpperTransformer : ICellTransformer
{
    public CellValue Transform(CellValue value, CellContext context) => new(value.Text?.ToUpperInvariant());
}

public class MinusOneFallback : ICellFallback<int>
{
    public int Fallback(CellValue value, CellContext context, Exception? exception) => -1;
}

public class Selected
{
    [ExcelColumnMatching(typeof(PrefixSelector), ConstructorArguments = ["Amount"])]
    public int Amount { get; set; }

    [ExcelColumnsMatching(typeof(PrefixSelector), ConstructorArguments = ["Q"])]
    public int[] Quarters { get; set; } = [];
}

public class PrefixSelector(string prefix) : IColumnSelector
{
    public IReadOnlyList<int> SelectColumns(ExcelSheet sheet)
        => sheet.Heading!.ColumnNames.Where(n => n.StartsWith(prefix, StringComparison.Ordinal)).Select(sheet.Heading.GetColumnIndex).ToArray();

    public string Describe(ExcelSheet sheet) => $"columns starting with \"{prefix}\"";
}
