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
    public string? Id { get; set; }
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
