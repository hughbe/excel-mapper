# ExcelMapper

A powerful, flexible .NET library for mapping Excel spreadsheet data to strongly-typed C# objects. ExcelMapper provides an intuitive fluent API with extensive customization options, robust type conversion, and comprehensive error handling.

![.NET Core](https://github.com/hughbe/excel-mapper/workflows/.NET%20Core/badge.svg)
[![Nuget](https://img.shields.io/nuget/v/ExcelDataReader.Mapping)](https://www.nuget.org/packages/ExcelDataReader.Mapping/)

Built on top of [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) for reliable Excel file parsing.

## Features

- ✨ **Automatic mapping** - Maps properties by convention with zero configuration
- 🎯 **Type-safe fluent API** - Strongly-typed mapping configuration using expressions
- 🔧 **Extensive customization** - Custom converters, transformers, and fallback strategies
- 📊 **Multiple mapping strategies** - One-to-one, many-to-one, collections, dictionaries
- 🏷️ **Attribute-based mapping** - Simple declarative mapping with attributes
- 🔄 **Flexible column selection** - By name, index, regex pattern, or custom predicate
- 🛡️ **Robust error handling** - Optional properties, default values, and custom fallbacks
- 🚀 **High performance** - Streaming API with lazy evaluation and caching
- 📦 **Rich type support** - Primitives, enums, DateTime, collections, nested objects, and more

## Quick Start

```csharp
using ExcelMapper;

// Define your model
public class Product
{
    public string Name { get; set; }
    public decimal Price { get; set; }
    public int Stock { get; set; }
}

// Read Excel data
using var importer = new ExcelImporter("products.xlsx");
var sheet = importer.ReadSheet();
var products = sheet.ReadRows<Product>().ToArray();
```

That's it! ExcelMapper automatically maps columns to properties by name.

## Table of Contents

- [Installation](#installation)
- [Basic Usage](#basic-usage)
- [Reading Workbooks](#reading-workbooks)
- [Reading Sheets](#reading-sheets)
- [Reading Rows](#reading-rows)
- [Mapping Strategies](#mapping-strategies)
  - [Automatic Mapping](#automatic-mapping)
  - [Attribute-Based Mapping](#attribute-based-mapping)
  - [Fluent API Mapping](#fluent-api-mapping)
- [Advanced Features](#advanced-features)
  - [Collections and Arrays](#collections-and-arrays)
  - [Dictionaries](#dictionaries)
  - [Nested Objects](#nested-objects)
  - [Enums](#enums)
  - [Custom Converters](#custom-converters)
- [Error Handling](#error-handling)
- [Performance Tips](#performance-tips)

## Installation

```bash
dotnet add package ExcelDataReader.Mapping
```

## Basic Usage

### Simple Example

| Name          | Email              | Age |
|---------------|--------------------|-----|
| John Smith    | john@example.com   | 32  |
| Jane Doe      | jane@example.com   | 28  |

```csharp
using ExcelMapper;

public class Person
{
    public string Name { get; set; }
    public string Email { get; set; }
    public int Age { get; set; }
}

using var importer = new ExcelImporter("people.xlsx");
var sheet = importer.ReadSheet();
var people = sheet.ReadRows<Person>().ToArray();

Console.WriteLine(people[0].Name);  // John Smith
Console.WriteLine(people[1].Age);   // 28
```

## Reading Workbooks

Create an `ExcelImporter` to read Excel or CSV files:

```csharp
// From file path
using var importer = new ExcelImporter("data.xlsx");

// From stream
using var stream = File.OpenRead("data.xlsx");
using var importer = new ExcelImporter(stream);

// CSV file
using var importer = new ExcelImporter("data.csv", ExcelImporterFileType.Csv);

// From existing IExcelDataReader
using var reader = ExcelReaderFactory.CreateReader(stream);
using var importer = new ExcelImporter(reader);
```

## Reading Sheets

### Read All Sheets

```csharp
foreach (var sheet in importer.ReadSheets())
{
    Console.WriteLine($"Sheet: {sheet.Name}");
    Console.WriteLine($"Visibility: {sheet.Visibility}");
    Console.WriteLine($"Columns: {sheet.NumberOfColumns}");
}
```

### Read Sheets Sequentially

```csharp
// Throws if no more sheets
var sheet1 = importer.ReadSheet();

// Returns false if no more sheets
if (importer.TryReadSheet(out var sheet2))
{
    // Process sheet2
}
```

### Read Sheet by Name

```csharp
// Throws if sheet doesn't exist
var sheet = importer.ReadSheet("Sales Data");

// Returns false if sheet doesn't exist
if (importer.TryReadSheet("Sales Data", out var salesSheet))
{
    // Process sheet
}
```

### Read Sheet by Index

```csharp
// Throws if index is invalid
var sheet = importer.ReadSheet(0);  // First sheet

// Returns false if index is invalid
if (importer.TryReadSheet(1, out var secondSheet))
{
    // Process sheet
}
```

## Reading Rows

### Read All Rows

```csharp
// Lazy evaluation - rows are read as you iterate
var rows = sheet.ReadRows<Product>();

// Or materialize to array
var products = sheet.ReadRows<Product>().ToArray();
```

### Read Specific Range

```csharp
// Read 10 rows starting from row index 5 (after header)
var rows = sheet.ReadRows<Product>(startIndex: 5, count: 10);
```

### Read Rows Sequentially

```csharp
// Throws if no more rows
var row1 = sheet.ReadRow<Product>();

// Returns false if no more rows
if (sheet.TryReadRow<Product>(out var row2))
{
    // Process row2
}
```

### Skip Blank Lines

```csharp
// Enable blank line skipping (off by default for performance)
importer.Configuration.SkipBlankLines = true;

var rows = sheet.ReadRows<Product>();
```

## Mapping Strategies

### Automatic Mapping

ExcelMapper automatically maps public properties and fields by matching column names (case-insensitive by default).
**Example:**

| Name          | Location         | Attendance | Date       | Link                  | Revenue | Cause   |
|---------------|------------------|------------|------------|-----------------------|---------|---------|
| Pub Quiz      | The Blue Anchor  | 20         | 2017-07-18 | http://eventbrite.com | 100.2   | Charity |
| Live Music    | The Raven        | 15         | 2017-07-17 | http://example.com    | 105.6   | Profit  |

```csharp
public enum EventCause { Profit, Charity }

public class Event
{
    public string Name { get; set; }
    public string Location { get; set; }
    public int Attendance { get; set; }
    public DateTime Date { get; set; }
    public Uri Link { get; set; }
    public EventCause Cause { get; set; }
}

using var importer = new ExcelImporter("events.xlsx");
var sheet = importer.ReadSheet();
var events = sheet.ReadRows<Event>().ToArray();

Console.WriteLine(events[0].Name);     // Pub Quiz
Console.WriteLine(events[0].Cause);    // Charity
Console.WriteLine(events[1].Revenue);  // 105.6
```

### Attribute-Based Mapping

Use attributes to declaratively configure mapping behavior.

#### Column Name Mapping

Map properties to columns with different names:

| Full Name      | #Age |
|----------------|------|
| Donald Trump   | 73   |
| Barack Obama   | 58   |

```csharp
public class President
{
    [ExcelColumnName("Full Name")]
    public string Name { get; set; }

    [ExcelColumnName("#Age")]
    public int Age { get; set; }
}

var presidents = sheet.ReadRows<President>().ToArray();
Console.WriteLine(presidents[0].Name);  // Donald Trump
Console.WriteLine(presidents[1].Age);   // 58
```

#### Multiple Column Name Variants

Try multiple column names in order of preference:

```csharp
public class President
{
    public string Name { get; set; }

    // Try these column names in order
    [ExcelColumnNames("Age", "#Age", "Years")]
    public int Age { get; set; }

    // Or use multiple attributes
    [ExcelColumnName("Party")]
    [ExcelColumnName("Political Party")]
    public string PoliticalParty { get; set; }
}
```

#### Pattern Matching

Match columns using regex patterns:

```csharp
public class Pub
{
    public string Name { get; set; }

    // Match columns like "2024 Attendance", "2025 Projected Attendance"
    [ExcelColumnMatching(@"\d{4}.*Attendance", RegexOptions.IgnoreCase)]
    public int Attendance { get; set; }
}
```

#### Column Index Mapping

Map by zero-based column index (useful for sheets without headers):

|                |    |
|----------------|----|
| Donald Trump   | 73 |
| Barack Obama   | 58 |

```csharp
public class President
{
    [ExcelColumnIndex(0)]
    public string Name { get; set; }

    [ExcelColumnIndex(1)]
    public int Age { get; set; }
}

var sheet = importer.ReadSheet();
sheet.HasHeading = false;  // No header row
var presidents = sheet.ReadRows<President>().ToArray();
```

#### Multiple Index Variants

```csharp
public class Data
{
    // Try column index 2, then 1, then 0
    [ExcelColumnIndices(2, 1, 0)]
    public string Value { get; set; }
}
```

#### Optional Properties

Skip properties if columns are missing:

```csharp
public class President
{
    public string Name { get; set; }

    [ExcelOptional]
    public int? Age { get; set; }  // Won't throw if column missing
}
```

#### Default Values

Provide default values for empty cells:
| Name         | Age |
|--------------|-----|
| Donald Trump |     |
| Barack Obama | 58  |

```csharp
public class President
{
    public string Name { get; set; }

    [ExcelDefaultValue(-1)]
    public int Age { get; set; }  // -1 if cell is empty
}
```

#### Ignore Properties

Exclude properties from mapping:

```csharp
public class President
{
    public string Name { get; set; }

    [ExcelIgnore]
    public int Age { get; set; }  // Never mapped from Excel

    [ExcelIgnore]
    public DateTime CreatedAt { get; set; }  // Computed property
}
```

#### Preserve Formatting

Read formatted string values instead of raw values:

| ID    | Price  |
|-------|--------|
| 00123 | $45.99 |
| 00456 | $12.50 |

```csharp
public class Product
{
    [ExcelPreserveFormatting]
    public string ID { get; set; }    // "00123" with leading zeros

    [ExcelPreserveFormatting]
    public string Price { get; set; }  // "$45.99" with currency symbol
}
```

### Fluent API Mapping

For complex scenarios, use fluent mapping with `ExcelClassMap<T>`:

```csharp
public class ProductMap : ExcelClassMap<Product>
{
    public ProductMap()
    {
        Map(p => p.Name)
            .WithColumnName("Product Name");

        Map(p => p.Price)
            .WithColumnIndex(2);

        Map(p => p.Category)
            .WithColumnNames("Category", "Type", "Classification")
            .MakeOptional();
    }
}

// Register the map
importer.Configuration.RegisterClassMap<ProductMap>();

var products = sheet.ReadRows<Product>();
```

#### Fluent Mapping Options

The fluent API provides extensive configuration options:

**Column Selection:**
- `.WithColumnName("Column Name")` - Map to specific column by name
- `.WithColumnIndex(0)` - Map to specific column by zero-based index
- `.WithColumnNames("Name1", "Name2")` - Try multiple column names in order
- `.WithColumnIndices(0, 1, 2)` - Try multiple indices in order
- `.WithColumnNameMatching(name => name.Contains("Total"))` - Use predicate
- `.WithColumnMatching(matcher)` - Use custom `IExcelColumnMatcher`

**Behavior:**
- `.MakeOptional()` - Don't throw if column is missing
- `.WithEmptyFallback(value)` - Use default value if cell is empty
- `.WithInvalidFallback(value)` - Use default value if conversion fails
- `.WithValueFallback(value)` - Use default value for both empty and invalid

**Advanced:**
- `.WithConverter(value => ...)` - Custom conversion delegate
- `.WithDateFormats("yyyy-MM-dd", "dd/MM/yyyy")` - Parse dates with specific formats
- `.WithMapping(dictionary)` - Map string values to enum/object values
- `.WithElementMap(...)` - Configure element pipeline for collections

#### Complete Fluent Example
```csharp
public enum MaritalStatus { Married, Divorced, Single }

public class President
{
    public string Name { get; set; }
    public MaritalStatus Status { get; set; }
    public int Children { get; set; }
    public float ApprovalRating { get; set; }
    public DateTime DateOfBirth { get; set; }
    public string Party { get; set; }
}

public class PresidentMap : ExcelClassMap<President>
{
    public PresidentMap()
    {
        Map(p => p.Name);

        // Map misspelled column and string values
        Map(p => p.Status)
            .WithColumnName("Marrital Status")
            .WithMapping(new Dictionary<string, MaritalStatus>
            {
                { "Twice Married", MaritalStatus.Married }
            });

        // Map by index
        Map(p => p.Children)
            .WithColumnIndex(2);

        // Custom converter
        Map(p => p.ApprovalRating)
            .WithColumnName("Approval Rating (%)")
            .WithConverter(value => float.Parse(value.TrimEnd('%')) / 100f);

        // Date parsing with multiple formats
        Map(p => p.DateOfBirth)
            .WithDateFormats("yyyy-MM-dd", "dd/MM/yyyy");

        // Try multiple column names
        Map(p => p.Party)
            .WithColumnNames("Political Party", "Party", "Affiliation");
    }
}

// Register and use
importer.Configuration.RegisterClassMap<PresidentMap>();
var presidents = sheet.ReadRows<President>();
```

## Error Handling

### Nullable Types and Fallbacks

By default:
- Nullable types are set to `null` for empty cells
- Non-nullable types throw `ExcelMappingException` for empty/invalid cells

Configure fallback behavior:

| Name         | Status  | Children | DateOfBirth |
|--------------|---------|----------|-------------|
| Donald Trump | invalid | invalid  | invalid     |
| Barack Obama |         |          |             |

```csharp
public enum MaritalStatus { Married, Single, Invalid, Unknown }

public class President
{
    public string Name { get; set; }
    public MaritalStatus Status { get; set; }
    public int? Children { get; set; }
    public DateTime? DateOfBirth { get; set; }
}

public class PresidentMap : ExcelClassMap<President>
{
    public PresidentMap()
    {
        Map(p => p.Name);

        Map(p => p.Status)
            .WithEmptyFallback(MaritalStatus.Unknown)     // Empty cells
            .WithInvalidFallback(MaritalStatus.Invalid);  // Invalid values

        Map(p => p.Children)
            .WithInvalidFallback(-1);  // Can't parse as int

        Map(p => p.DateOfBirth)
            .WithInvalidFallback(null);  // Can't parse as DateTime
    }
}

importer.Configuration.RegisterClassMap<PresidentMap>();
var presidents = sheet.ReadRows<President>();
```

## Advanced Features

### Enums

Parse string values to enums (case-sensitive by default):

| Name         | Status   |
|--------------|----------|
| Donald Trump | Married  |
| Barack Obama | married  |
| Joe Biden    | DIVORCED |

```csharp
public enum MaritalStatus { Married, Divorced, Single }

public class President
{
    public string Name { get; set; }
    public MaritalStatus Status { get; set; }
}

// Case-insensitive enum parsing
public class PresidentMap : ExcelClassMap<President>
{
    public PresidentMap()
    {
        Map(p => p.Name);
        Map(p => p.Status, ignoreCase: true);  // Handles "married", "MARRIED", etc.
    }
}
```

### Collections and Arrays

ExcelMapper supports multiple strategies for mapping collections.

#### Split Single Cell

By default, splits cell value by comma:

| Name         | Tags                     |
|--------------|--------------------------|
| Barack Obama | President,Democrat,2000s |

```csharp
public class Person
{
    public string Name { get; set; }
    public string[] Tags { get; set; }  // Auto-split by comma
}
```

#### Multiple Columns by Name

```csharp
public class Pub
{
    public string Name { get; set; }

    [ExcelColumnNames("Drink1", "Drink2", "Drink3")]
    public string[] Drinks { get; set; }
}
```

#### Multiple Columns by Index

```csharp
public class Pub
{
    public string Name { get; set; }

    [ExcelColumnIndices(1, 2, 3)]
    public string[] Drinks { get; set; }
}
```

#### Multiple Columns by Pattern

```csharp
public class Pub
{
    public string Name { get; set; }

    [ExcelColumnsMatching(@"Drink\d+", RegexOptions.IgnoreCase)]
    public string[] Drinks { get; set; }
}
```

#### Fluent Collection Mapping

```csharp
public class President
{
    public string Name { get; set; }
    public string[] Children { get; set; }
    public DateTime[] Elections { get; set; }
}

public class PresidentMap : ExcelClassMap<President>
{
    public PresidentMap()
    {
        Map(p => p.Name);

        // Split by comma (default)
        Map(p => p.Children)
            .WithColumnName("Children Names");

        // Read multiple columns
        Map(p => p.Elections)
            .WithColumnNames("First Election", "Second Election")
            .WithElementMap(m => m
                .WithDateFormats("yyyy-MM-dd", "dd/MM/yyyy")
            );
    }
}
```

#### Supported Collection Types

- `T[]` - Arrays
- `List<T>`, `IList<T>`, `ICollection<T>`, `IEnumerable<T>`
- `HashSet<T>`, `ISet<T>`
- `FrozenSet<T>`, `ImmutableHashSet<T>`
- And more...

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
- `FrozenDictionary<TKey, TValue>`
- `ImmutableDictionary<TKey, TValue>`
- Keys are derived from column names

### Nested Objects

Map nested properties to Excel columns:


| Name         | Elected    | Votes |
|--------------|------------|-------|
| Barack Obama | 2008-11-04 | 365   |

```csharp
public class Election
{
    public DateTime Date { get; set; }
    public int Votes { get; set; }
}

public class President
{
    public string Name { get; set; }
    public Election ElectionInfo { get; set; }
}

public class PresidentMap : ExcelClassMap<President>
{
    public PresidentMap()
    {
        Map(p => p.Name);

        // Map nested properties
        Map(p => p.ElectionInfo.Date)
            .WithColumnName("Elected");

        Map(p => p.ElectionInfo.Votes);
    }
}

importer.Configuration.RegisterClassMap<PresidentMap>();
var presidents = sheet.ReadRows<President>();
```

### Custom Converters

Create custom type conversions:

```csharp
public class ProductMap : ExcelClassMap<Product>
{
    public ProductMap()
    {
        Map(p => p.Price)
            .WithConverter(value => 
            {
                // Remove currency symbol and parse
                var cleaned = value.Replace("$", "").Replace(",", "");
                return decimal.Parse(cleaned);
            });

        Map(p => p.Available)
            .WithConverter(value => value.ToLower() switch
            {
                "yes" => true,
                "y" => true,
                "no" => false,
                "n" => false,
                _ => false
            });
    }
}
```

## Special Scenarios

### Sheets Without Headers

Disable header row and use column indices:

|           |                  |
|-----------|------------------|
| Pub Quiz  | The Blue Anchor  |
| Live Music| The Raven        |

```csharp
public class Event
{
    public string Name { get; set; }
    public string Location { get; set; }
}

public class EventMap : ExcelClassMap<Event>
{
    public EventMap()
    {
        Map(e => e.Name).WithColumnIndex(0);
        Map(e => e.Location).WithColumnIndex(1);
    }
}

using var importer = new ExcelImporter("events.xlsx");
importer.Configuration.RegisterClassMap<EventMap>();

var sheet = importer.ReadSheet();
sheet.HasHeading = false;  // Disable header row

var events = sheet.ReadRows<Event>();
```

### Headers Not in First Row

Skip rows before the header:

|               |          |
|---------------|----------|
| Report Title  |          |
|               |          |
| Name          | Location |
| Pub Quiz      | Downtown |
| Live Music    | Uptown   |

```csharp
public class Event
{
    public string Name { get; set; }
    public string Location { get; set; }
}

using var importer = new ExcelImporter("events.xlsx");
var sheet = importer.ReadSheet();
sheet.HeadingIndex = 2;  // Header is on row 3 (zero-based index 2)

var events = sheet.ReadRows<Event>();
```

## Performance Tips

1. **Use streaming**: `ReadRows<T>()` uses lazy evaluation - don't materialize unnecessarily
   ```csharp
   // Good - processes one at a time
   foreach (var product in sheet.ReadRows<Product>())
   {
       ProcessProduct(product);
   }

   // Avoid - loads everything into memory
   var allProducts = sheet.ReadRows<Product>().ToList();
   ```

2. **Register maps once**: Class maps are cached per type
   ```csharp
   importer.Configuration.RegisterClassMap<ProductMap>();
   ```

3. **Disable blank line skipping**: Off by default for performance
   ```csharp
   importer.Configuration.SkipBlankLines = false;  // Default
   ```

4. **Use column indices for headerless sheets**: Faster than column name lookup
   ```csharp
   Map(p => p.Name).WithColumnIndex(0);  // Faster
   Map(p => p.Name).WithColumnName("Name");  // Requires lookup
   ```

## Supported Types

### Primitive Types
- Numeric: `int`, `long`, `double`, `decimal`, `float`, `byte`, `short`, `uint`, `ulong`, `ushort`, `sbyte`
- Text: `string`, `char`
- Other: `bool`, `DateTime`, `Guid`, `Uri`
- Enums (with optional case-insensitive parsing)
- Nullable versions of all value types

### Collection Types
- `T[]`, `List<T>`, `IList<T>`, `ICollection<T>`, `IEnumerable<T>`
- `HashSet<T>`, `ISet<T>`, `FrozenSet<T>`, `ImmutableHashSet<T>`
- `Dictionary<TKey, TValue>`, `IDictionary<TKey, TValue>`
- `FrozenDictionary<TKey, TValue>`, `ImmutableDictionary<TKey, TValue>`

### Complex Types
- Classes with public properties/fields
- Record types
- Nested objects
- `ExpandoObject` for dynamic scenarios

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

Built on top of [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) for robust Excel file parsing.
