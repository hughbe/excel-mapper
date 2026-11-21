# Types and Advanced Mapping

Supported types, enums, collections, dictionaries, nested objects, custom converters and transformers. See the [documentation index](../README.md#documentation) for other topics.

## Advanced Features

### Enums

Parse string values to enums (case-sensitive by default):

| Name          | Status   |
|---------------|----------|
| Alice Johnson | FullTime |
| Bob Smith     | fulltime |
| Carol White   | PARTTIME |

```csharp
public enum EmploymentStatus { FullTime, PartTime, Contract }

public class Employee
{
    public string Name { get; set; }
    public EmploymentStatus Status { get; set; }
}

// Case-insensitive enum parsing
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name);
        Map(e => e.Status, ignoreCase: true);  // Handles "fulltime", "FULLTIME", etc.
    }
}
```

#### Mapping Enums with DescriptionAttribute

ExcelMapper automatically supports mapping enum values using their `[Description]` attribute values. This is useful when your Excel data contains user-friendly text that differs from the enum names:

| Status              |
|---------------------|
| Full Time Employee  |
| Part Time Employee  |
| Contract Worker     |

```csharp
using System.ComponentModel;

public enum EmploymentStatus
{
    [Description("Full Time Employee")]
    FullTime,
    
    [Description("Part Time Employee")]
    PartTime,
    
    [Description("Contract Worker")]
    Contract
}

public class Employee
{
    public EmploymentStatus Status { get; set; }
}
```

The DescriptionAttribute mapping works automatically with:
- Direct enum properties (automatic mapping)
- Explicit property mapping via `RegisterClassMap`
- Nullable enums
- Both standard enum name matching and description matching (description takes precedence when both match)

**Important Notes:**
- Description matching is **case-sensitive**
- Empty or null descriptions are ignored
- If a description is not found, the mapper falls back to standard enum name matching
- Enum values without a `[Description]` attribute can still be mapped using their standard names


    [Fact]
    public void ReadRow_AutoMappedUriKindAttributeRelativeOrAbsolute_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value (absolute URI).
        var row1 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Uri);

        // Empty cell value.
        var row2 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Null(row2.Uri);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriKindRelativeOrAbsoluteClass>());
    }

    private class UriKindRelativeOrAbsoluteClass
    {
        [ExcelUri(UriKind.RelativeOrAbsolute)]
        public Uri Uri { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_DefaultMappedUriKindAttributeRelativeOrAbsolute_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Uris.xlsx");
        importer.Configuration.RegisterClassMap<UriKindRelativeOrAbsoluteClass>(c =>
        {
            c.Map(u => u.Uri);
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        // Valid cell value (absolute URI).
        var row1 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Equal(new Uri("http://google.com"), row1.Uri);

        // Empty cell value.
        var row2 = sheet.ReadRow<UriKindRelativeOrAbsoluteClass>();
        Assert.Null(row2.Uri);

        // Invalid cell value.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<UriKindRelativeOrAbsoluteClass>());
    }
### Uri Mapping

ExcelMapper supports mapping cell values to `Uri` objects with configurable URI kind validation:

| Website                   | API Endpoint         | Relative Path    |
|---------------------------|----------------------|------------------|
| https://example.com       | /api/v1/users        | ../resources     |
| https://www.contoso.com   | /api/v1/products     | ./images/logo.png|

```csharp
public class Resource
{
    public Uri Website { get; set; }         // RelativeOrAbsolute URI
    public Uri ApiEndpoint { get; set; }     // Relative URI
    public Uri RelativePath { get; set; }    // Relative or absolute
}
```

#### Controlling URI Kind

By default, URIs accept any valid format (absolute, relative, or both). You can specify URI kind requirements using either attributes or the fluent API:

**Using Attributes:**

```csharp
public class Resource
{
    // Require absolute URIs only (e.g., https://example.com)
    [ExcelUri(UriKind.RelativeOrAbsolute)]
    public Uri Website { get; set; }
    
    // Require relative URIs only (e.g., /api/v1/users)
    [ExcelUri(UriKind.Relative)]
    public Uri ApiEndpoint { get; set; }
    
    // Allow both absolute and relative URIs (default)
    [ExcelUri(UriKind.RelativeOrRelativeOrAbsolute)]
    public Uri RelativePath { get; set; }
}

var resources = sheet.ReadRows<Resource>();
```

**Using Fluent API:**

```csharp
public class ResourceMap : ExcelClassMap<Resource>
{
    public ResourceMap()
    {
        // Require absolute URIs only (e.g., https://example.com)
        Map(r => r.Website)
            .WithUriKind(UriKind.RelativeOrAbsolute);
            
        // Require relative URIs only (e.g., /api/v1/users)
        Map(r => r.ApiEndpoint)
            .WithUriKind(UriKind.Relative);
            
        // Allow both absolute and relative URIs (default)
        Map(r => r.RelativePath)
            .WithUriKind(UriKind.RelativeOrRelativeOrAbsolute);
    }
}

importer.Configuration.RegisterClassMap<ResourceMap>();
var resources = sheet.ReadRows<Resource>();
```

**Available UriKind Options:**
- `UriKind.RelativeOrAbsolute` - Must be absolute URI (e.g., `https://example.com`)
- `UriKind.Relative` - Must be relative URI (e.g., `/api/users` or `../path`)
- `UriKind.RelativeOrRelativeOrAbsolute` (default) - Can be either absolute or relative

**Error Handling:**

```csharp
public class ResourceMap : ExcelClassMap<Resource>
{
    public ResourceMap()
    {
        Map(r => r.Website)
            .WithUriKind(UriKind.RelativeOrAbsolute)
            .WithInvalidFallback(null);  // Use null for invalid URIs
    }
}
```

**Important Notes:**
- **Hyperlink formulas are not supported** due to limitations in the underlying ExcelDataReader library. Only the displayed text value in cells is read, not the hyperlink target. If you need to read hyperlink targets, consider using a different library or extracting them separately.
- Invalid URIs will throw an exception unless you provide a fallback value
- Empty cells are treated as null for nullable `Uri?` properties
- URI validation is performed using .NET's `Uri.TryCreate` method

### Collections and Arrays

ExcelMapper supports multiple strategies for mapping collections.

#### Split Single Cell

By default, splits cell value by comma:

| Name          | Skills                           |
|---------------|----------------------------------|
| Alice Johnson | C#,Python,SQL                    |
| Bob Smith     | Java,JavaScript,Docker,Kubernetes |

```csharp
public class Employee
{
    public string Name { get; set; }
    public string[] Skills { get; set; }  // Auto-split by comma
}
```

Customize the separator using attributes or fluent API:

```csharp
// Using attributes
public class Employee
{
    public string Name { get; set; }
    
    // Split by semicolon instead of comma
    [ExcelSeparators(';')]
    public string[] Skills { get; set; }
    
    // Split by multiple separators (pipe or comma)
    [ExcelSeparators('|', ',')]
    public string[] Tags { get; set; }
    
    // String separators also supported
    [ExcelSeparators(";", "||")]
    public string[] Categories { get; set; }
}

// Or using fluent API
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name);
        
        // Split by semicolon instead of comma
        Map(e => e.Skills)
            .WithSeparators(';');
            
        // Split by multiple separators (pipe or comma)
        Map(e => e.Tags)
            .WithSeparators('|', ',');
    }
}
```

#### Multiple Columns by Name

```csharp
public class Employee
{
    public string Name { get; set; }

    [ExcelColumnNames("Review1", "Review2", "Review3")]
    public int[] Reviews { get; set; }
}
```

#### Multiple Columns by Index

```csharp
public class Employee
{
    public string Name { get; set; }

    [ExcelColumnIndices(1, 2, 3)]
    public int[] QuarterlyScores { get; set; }
}
```

#### Multiple Columns by Pattern

```csharp
public class Employee
{
    public string Name { get; set; }

    [ExcelColumnsMatching(@"Q\d+.*Score", RegexOptions.IgnoreCase)]
    public int[] QuarterlyScores { get; set; }
}
```

The fluent API accepts a predicate or a regular expression:

```csharp
Map(e => e.QuarterlyScores).WithColumnsMatching(name => name.StartsWith("Q"));
Map(e => e.QuarterlyScores).WithColumnsMatching(new Regex(@"Q\d+.*Score", RegexOptions.IgnoreCase));
```

#### Fluent Collection Mapping

```csharp
public class Employee
{
    public string Name { get; set; }
    public string[] Skills { get; set; }
    public DateTime[] Certifications { get; set; }
    public int[] Scores { get; set; }
}

public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name);

        // Split by comma (default)
        Map(e => e.Skills)
            .WithColumnName("Skills");

        // Read multiple columns with custom element mapping
        Map(e => e.Certifications)
            .WithColumnNames("Certification Date 1", "Certification Date 2")
            .WithElementMap(m => m
                .WithFormats("yyyy-MM-dd", "dd/MM/yyyy")
                .WithInvalidFallback(DateTime.MinValue)
            );
            
        // Configure element conversion for split values
        Map(e => e.Scores)
            .WithColumnName("Quarterly Scores")
            .WithSeparators(';')
            .WithElementMap(m => m
                .WithInvalidFallback(-1)  // Handle non-numeric values
            );
    }
}
```

The `.WithElementMap()` method allows you to configure how individual elements in a collection are parsed, including:
- Custom formats (`.WithFormats()`)
- Fallback values (`.WithEmptyFallback()`, `.WithInvalidFallback()`)
- Custom converters (`.WithConverter()`)
- Value mappings (`.WithMapping()`)

#### Supported Collection Types

- **Arrays**: `T[]`
- **Lists**: `List<T>`, `IList<T>`, `ICollection<T>`, `IEnumerable<T>`
- **Sets**: `HashSet<T>`, `ISet<T>`, `FrozenSet<T>` (.NET 8+), `ImmutableHashSet<T>`, `ImmutableSortedSet<T>`
- **Immutable Collections**: `ImmutableArray<T>`, `ImmutableList<T>`
- **Observable Collections**: `ObservableCollection<T>`
- **Custom collections** with `Add(T)` method and parameterless constructor

### Dictionaries

Map multiple columns to dictionary properties.

#### All Columns to Dictionary

```csharp
// Maps ALL columns to dictionary
var rows = sheet.ReadRows<Dictionary<string, string>>();
Console.WriteLine(rows[0]["Name"]);
Console.WriteLine(rows[0]["Age"]);
```

#### Dictionary Property

```csharp
public class Record
{
    public Dictionary<string, string> Values { get; set; }
}

public class RecordMap : ExcelClassMap<Record>
{
    public RecordMap()
    {
        // Map all columns
        Map(r => r.Values);

        // Or specific columns
        Map(r => r.Values)
            .WithColumnNames("Column1", "Column2", "Column3");
    }
}
```

#### Supported Dictionary Types

- `Dictionary<TKey, TValue>`, `IDictionary<TKey, TValue>`
- `FrozenDictionary<TKey, TValue>` (.NET 8+)
- `ImmutableDictionary<TKey, TValue>`, `ImmutableSortedDictionary<TKey, TValue>`
- Keys are derived from column names
- Values can be any supported type

### Nested Objects

Map nested properties to Excel columns:

| Name          | HireDate   | Department | Location |
|---------------|------------|------------|----------|
| Alice Johnson | 2020-03-15 | Engineering| Seattle  |
| Bob Smith     | 2019-07-22 | Marketing  | New York |

```csharp
public class DepartmentInfo
{
    public string Name { get; set; }
    public string Location { get; set; }
}

public class Employee
{
    public string Name { get; set; }
    public DateTime HireDate { get; set; }
    public DepartmentInfo Department { get; set; }
}

public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name);
        Map(e => e.HireDate);

        // Map nested properties
        Map(e => e.Department.Name)
            .WithColumnName("Department");

        Map(e => e.Department.Location);
    }
}

importer.Configuration.RegisterClassMap<EmployeeMap>();
var employees = sheet.ReadRows<Employee>();
```

**Circular Reference Detection:**

ExcelMapper automatically detects and prevents circular references during auto-mapping:

```csharp
public class Person
{
    public string Name { get; set; }
    public Person Parent { get; set; }  // Would cause infinite recursion
}

// This will throw ExcelMappingException with a clear error message
var people = sheet.ReadRows<Person>();
// Exception: "Circular reference detected: type 'Person' references itself 
//             through its members. Consider applying the ExcelIgnore 
//             attribute to break the cycle."
```

**Solution - use `[ExcelIgnore]`:**

```csharp
public class Person
{
    public string Name { get; set; }
    
    [ExcelIgnore]  // Break the circular reference
    public Person Parent { get; set; }
}
```

### Custom Converters

Create custom type conversions:

```csharp
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Salary)
            .WithConverter(value => 
            {
                // Remove currency symbol and parse
                var cleaned = value.Replace("$", "").Replace(",", "");
                return decimal.Parse(cleaned);
            });

        Map(e => e.Active)
            .WithConverter(value => value.ToLower() switch
            {
                "yes" => true,
                "y" => true,
                "active" => true,
                "no" => false,
                "n" => false,
                "inactive" => false,
                _ => false
            });
    }
}
```

### Custom Transformers and Mappers

For advanced scenarios, implement `ICellTransformer` or `ICellMapper`:

**ICellTransformer** - Transforms string values before mapping:

```csharp
public class UpperCaseTransformer : ICellTransformer
{
    public string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
    {
        return readResult.StringValue?.ToUpperInvariant();
    }
}

public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name)
            .WithTransformers(new UpperCaseTransformer());
            
        // Or use the built-in trim transformer
        Map(e => e.Department)
            .WithTrim();  // Convenience method for TrimStringCellTransformer
    }
}
```

**ICellMapper** - Custom type conversion logic:

```csharp
public class PhoneNumberMapper : ICellMapper
{
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var value = readResult.StringValue;
        if (string.IsNullOrWhiteSpace(value))
        {
            return CellMapperResult.Empty();
        }

        try
        {
            // Remove formatting and validate
            var cleaned = new string(value.Where(char.IsDigit).ToArray());
            if (cleaned.Length != 10)
            {
                return CellMapperResult.Invalid(
                    new FormatException("Phone number must be 10 digits"));
            }
            
            return CellMapperResult.Success(cleaned);
        }
        catch (Exception ex)
        {
            return CellMapperResult.Invalid(ex);
        }
    }
}

public class ContactMap : ExcelClassMap<Contact>
{
    public ContactMap()
    {
        Map(e => e.PhoneNumber)
            .WithMappers(new PhoneNumberMapper());
    }
}
```

**Chaining Transformers and Mappers:**

```csharp
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name)
            .WithTransformers(
                new TrimStringCellTransformer(),
                new UpperCaseTransformer()
            )
            .WithMappers(new CustomStringMapper());
    }
}
```

## Supported Types

### Primitive Types
- **Numeric**: `int`, `long`, `double`, `decimal`, `float`, `byte`, `short`, `uint`, `ulong`, `ushort`, `sbyte`
- **Extended Numeric**: `Int128`, `UInt128`, `BigInteger`, `Half`, `nint` (IntPtr), `nuint` (UIntPtr), `Complex`
- **Text**: `string`, `char`
- **Other**: `bool` (Excel booleans, `true`/`false`, `1`/`0` and `yes`/`no`, case-insensitive), `Guid`, `Uri`, `Version`
- **Enums**: With optional case-insensitive parsing
- **Nullable versions** of all value types
- **Any type implementing `IParsable<T>`** (.NET 7+) - automatically supported
- **Any type implementing `IConvertible`** - automatically supported

> **Note:** Fractional numeric cells (such as `2.7`) mapped to integer types are invalid rather than truncated or rounded. See [Fractional Numbers Mapped to Integer Properties](errors.md#fractional-numbers-mapped-to-integer-properties).

> **Note:** `object` and `IConvertible` members, and `object` values in dictionaries, `ExpandoObject` and non-generic collections such as `ArrayList`, receive the value as Excel stores it: a `double` for numbers, a `DateTime` for dates, a `bool` for booleans and a `string` for text. With `[ExcelPreserveFormatting]` they receive the displayed text instead. CSV cells are always text. Collections that only accept one type, such as `StringCollection`, convert values to that type.

### Collection Types
- `T[]`, `List<T>`, `IList<T>`, `ICollection<T>`, `IEnumerable<T>`
- `HashSet<T>`, `ISet<T>`, `FrozenSet<T>`, `ImmutableHashSet<T>`
- `ImmutableArray<T>`, `ImmutableList<T>`
- `Dictionary<TKey, TValue>`, `IDictionary<TKey, TValue>`
- `FrozenDictionary<TKey, TValue>`, `ImmutableDictionary<TKey, TValue>`

### Date and Time Types

ExcelMapper has comprehensive support for modern .NET date and time types:

- `DateTime` - Date and time
- `DateTimeOffset` - Date and time with an offset. Excel dates have no time zone, so by default they get the machine's local offset; use `[ExcelOffset("+00:00")]` or `.WithOffset(TimeSpan.Zero)` to choose one
- `DateOnly` - Date without time (available in .NET 6+)
- `TimeOnly` - Time without date (available in .NET 6+)
- `TimeSpan` - Duration/time interval

Cells that Excel stores as dates, times or durations are mapped directly. Excel stores dates as numbers (the number of days since 1899-12-30), so a number in a cell without a date format is also read as a date: `45355` maps to 2024-03-04 for `DateTime`, `DateTimeOffset` and `DateOnly`. For `TimeOnly` a number is a fraction of a day (`0.25` is 06:00), and for `TimeSpan` it is a number of days (`1.5` is 36 hours). A cell formatted as a time of day can be mapped to `TimeOnly` or `TimeSpan`.

Text cells support custom format parsing with `.WithFormats()` and culture-specific parsing with `.WithFormatProvider()`.

**Example:**

```csharp
using System.Globalization;

public class EventMap : ExcelClassMap<Event>
{
    public EventMap()
    {
        Map(e => e.EventDate)
            .WithFormats("yyyy-MM-dd", "dd/MM/yyyy", "MM/dd/yyyy");
            
        Map(e => e.StartTime)
            .WithFormats("HH:mm:ss", "hh:mm tt");
            
        Map(e => e.Duration)
            .WithFormats(@"hh\:mm\:ss", @"mm\:ss");
    }
}
```

**`DateTimeOffset` offsets:**

Excel stores dates without a time zone. When mapped to a `DateTimeOffset`, they get the machine's local offset by default, so the same file can give different results on different servers. Choose the offset instead:

```csharp
public class Event
{
    [ExcelOffset("+00:00")]          // UTC
    public DateTimeOffset StartsAt { get; set; }

    [ExcelOffset("-08:00")]
    public DateTimeOffset? EndsAt { get; set; }
}

// Or with the fluent API
Map(e => e.StartsAt).WithOffset(TimeSpan.Zero);
```

This applies to cells that Excel stores as dates (or as serial numbers). Text that includes an offset, parsed with a format such as `"yyyy-MM-ddTHH:mm:sszzz"`, keeps its own offset.

**Culture-Specific Parsing:**

Use `.WithFormatProvider()` to parse dates and times using specific cultures:

```csharp
public class EventMap : ExcelClassMap<Event>
{
    public EventMap()
    {
        var frenchCulture = new CultureInfo("fr-FR");
        
        // Parse French-formatted dates: "15/03/2024"
        Map(e => e.EventDate)
            .WithFormats("dd/MM/yyyy")
            .WithFormatProvider(frenchCulture);
    }
}
```

### Complex Types
- Classes with public properties/fields
- Record types (C# 9+) and classes or structs with get-only properties set through a constructor
- Nested objects
- `ExpandoObject` for dynamic scenarios

**Record Type Example:**

```csharp
// Records are created through their constructor (see Records and Constructors)
public record Employee(string Name, string Department, decimal Salary);

using var importer = new ExcelImporter("employees.xlsx");
var sheet = importer.ReadSheet();
var employees = sheet.ReadRows<Employee>().ToArray();
```

**ExpandoObject Example:**

```csharp
using System.Dynamic;

// Read rows as dynamic objects when column structure varies
using var importer = new ExcelImporter("data.xlsx");
var sheet = importer.ReadSheet();

foreach (dynamic row in sheet.ReadRows<ExpandoObject>())
{
    Console.WriteLine(row.ColumnName);  // Access properties dynamically
}

// Or use in a property
public class FlexibleData
{
    public string Id { get; set; }
    public ExpandoObject Metadata { get; set; }  // Maps all remaining columns
}
```
