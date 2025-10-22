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
public class Employee
{
    public string Name { get; set; }
    public string Department { get; set; }
    public decimal Salary { get; set; }
}

// Read Excel data
using var importer = new ExcelImporter("employees.xlsx");
var sheet = importer.ReadSheet();
var employees = sheet.ReadRows<Employee>().ToArray();
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
- [Special Scenarios](#special-scenarios)
  - [Sheets Without Headers](#sheets-without-headers)
  - [Headers Not in First Row](#headers-not-in-first-row)
- [Error Handling](#error-handling)
- [Performance Tips](#performance-tips)
- [Supported Types](#supported-types)
- [CSV Support](#csv-support)
- [Common Issues & Troubleshooting](#common-issues--troubleshooting)
- [Best Practices](#best-practices)
- [Contributing](#contributing)
- [License](#license)

## Installation

```bash
dotnet add package ExcelDataReader.Mapping
```

## Basic Usage

### Simple Example

| Name          | Department  | Salary  |
|---------------|-------------|---------|
| Alice Johnson | Engineering | 95000   |
| Bob Smith     | Marketing   | 78000   |

```csharp
using ExcelMapper;

public class Employee
{
    public string Name { get; set; }
    public string Department { get; set; }
    public decimal Salary { get; set; }
}

using var importer = new ExcelImporter("employees.xlsx");
var sheet = importer.ReadSheet();
var employees = sheet.ReadRows<Employee>().ToArray();

Console.WriteLine(employees[0].Name);      // Alice Johnson
Console.WriteLine(employees[1].Salary);    // 78000
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
var rows = sheet.ReadRows<Employee>();

// Or materialize to array
var employees = sheet.ReadRows<Employee>().ToArray();
```

### Read Specific Range

```csharp
// Read 10 rows starting from row index 5 (after header)
var rows = sheet.ReadRows<Employee>(startIndex: 5, count: 10);
```

### Read Rows Sequentially

```csharp
// Throws if no more rows
var row1 = sheet.ReadRow<Employee>();

// Returns false if no more rows
if (sheet.TryReadRow<Employee>(out var row2))
{
    // Process row2
}
```

### Skip Blank Lines

```csharp
// Enable blank line skipping (off by default for performance)
importer.Configuration.SkipBlankLines = true;

var rows = sheet.ReadRows<Employee>();
```

### Security: Column Count Limits

To protect against denial-of-service attacks from malicious Excel files with excessive columns, ExcelMapper enforces a maximum column limit per sheet:

```csharp
using var importer = new ExcelImporter("data.xlsx");

// Default limit is 10,000 columns (sufficient for most use cases)
Console.WriteLine(importer.Configuration.MaxColumnsPerSheet);  // 10000

// Adjust the limit if needed for legitimate large files
importer.Configuration.MaxColumnsPerSheet = 20000;

// Or disable the limit entirely (not recommended for untrusted files)
importer.Configuration.MaxColumnsPerSheet = int.MaxValue;
```

**Note:** Excel .xlsx files support up to 16,384 columns (XFD). If a sheet exceeds `MaxColumnsPerSheet`, an `ExcelMappingException` is thrown with a clear error message.

**Security Best Practices:**
- Keep the default limit (10,000) for untrusted/user-uploaded files
- Only increase the limit when you control the file source
- Files exceeding the limit will fail immediately before allocating excessive memory

## Mapping Strategies

### Automatic Mapping

ExcelMapper automatically maps public properties and fields by matching column names (case-insensitive by default).
**Example:**

| Name          | Department  | Position        | HireDate   | Salary | Active |
|---------------|-------------|-----------------|------------|--------|--------|
| Alice Johnson | Engineering | Senior Engineer | 2020-03-15 | 95000  | true   |
| Bob Smith     | Marketing   | Manager         | 2019-07-22 | 78000  | true   |

```csharp
public class Employee
{
    public string Name { get; set; }
    public string Department { get; set; }
    public string Position { get; set; }
    public DateTime HireDate { get; set; }
    public decimal Salary { get; set; }
    public bool Active { get; set; }
}

using var importer = new ExcelImporter("employees.xlsx");
var sheet = importer.ReadSheet();
var employees = sheet.ReadRows<Employee>().ToArray();

Console.WriteLine(employees[0].Name);       // Alice Johnson
Console.WriteLine(employees[0].Position);   // Senior Engineer
Console.WriteLine(employees[1].Salary);     // 78000
```

### Attribute-Based Mapping

Use attributes to declaratively configure mapping behavior.

#### Column Name Mapping

Map properties to columns with different names:

| Full Name      | #Age |
|----------------|------|
| Alice Johnson  | 32   |
| Bob Smith      | 45   |

```csharp
public class Employee
{
    [ExcelColumnName("Full Name")]
    public string Name { get; set; }

    [ExcelColumnName("#Age")]
    public int Age { get; set; }
}

var employees = sheet.ReadRows<Employee>().ToArray();
Console.WriteLine(employees[0].Name);  // Alice Johnson
Console.WriteLine(employees[1].Age);   // 45
```

#### Multiple Column Name Variants

Try multiple column names in order of preference:

```csharp
public class Employee
{
    public string Name { get; set; }

    // Try these column names in order
    [ExcelColumnNames("Age", "#Age", "Years")]
    public int Age { get; set; }

    // Or use multiple attributes
    [ExcelColumnName("Dept")]
    [ExcelColumnName("Department")]
    public string Department { get; set; }
}
```

#### Pattern Matching

Match columns using regex patterns or custom matchers:

```csharp
public class Employee
{
    public string Name { get; set; }

    // Match columns like "2024 Salary", "2025 Projected Salary"
    [ExcelColumnMatching(@"\d{4}.*Salary", RegexOptions.IgnoreCase)]
    public decimal Salary { get; set; }
}
```

For advanced matching logic, implement `IExcelColumnMatcher`:

```csharp
public class StartsWithMatcher : IExcelColumnMatcher
{
    private readonly string _prefix;

    public StartsWithMatcher(string prefix)
    {
        _prefix = prefix;
    }

    public bool IsMatch(string columnName) => columnName.StartsWith(_prefix);
}

public class Employee
{
    // Use custom matcher to match columns starting with "Bonus_"
    [ExcelColumnsMatching(typeof(StartsWithMatcher), ConstructorArguments = new object[] { "Bonus_" })]
    public decimal TotalBonus { get; set; }
}
```

#### Column Index Mapping

Map by zero-based column index (useful for sheets without headers):

|                |    |
|----------------|----|
| Alice Johnson  | 32 |
| Bob Smith      | 45 |

```csharp
public class Employee
{
    [ExcelColumnIndex(0)]
    public string Name { get; set; }

    [ExcelColumnIndex(1)]
    public int Age { get; set; }
}

var sheet = importer.ReadSheet();
sheet.HasHeading = false;  // No header row
var employees = sheet.ReadRows<Employee>().ToArray();
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
public class Employee
{
    public string Name { get; set; }

    [ExcelOptional]
    public int? Age { get; set; }  // Won't throw if column missing
}
```

#### Default Values

Provide default values for empty cells:
| Name          | Age |
|---------------|-----|
| Alice Johnson |     |
| Bob Smith     | 45  |

```csharp
public class Employee
{
    public string Name { get; set; }

    [ExcelDefaultValue(-1)]
    public int Age { get; set; }  // -1 if cell is empty
}
```

#### Ignore Properties

Exclude properties from mapping:

```csharp
public class Employee
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

| Employee ID | Salary   |
|-------------|----------|
| 00123       | $95,000  |
| 00456       | $78,000  |

```csharp
public class Employee
{
    [ExcelPreserveFormatting]
    public string EmployeeID { get; set; }    // "00123" with leading zeros

    [ExcelPreserveFormatting]
    public string Salary { get; set; }  // "$95,000" with currency symbol
}
```

#### Trim String Values

Automatically trim whitespace from string values:

| Name             |
|------------------|
|  Alice Johnson   |
|   Bob Smith      |

```csharp
public class Employee
{
    [ExcelTrimString]
    public string Name { get; set; }  // "Alice Johnson", "Bob Smith" (trimmed)
}
```

Or use the fluent API:

```csharp
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name).WithTrim();
    }
}
```

#### Invalid Value Fallback

Provide a fallback value when cell value cannot be parsed:

| Name          | Age    |
|---------------|--------|
| Alice Johnson | 32     |
| Bob Smith     | N/A    |

```csharp
public class Employee
{
    public string Name { get; set; }

    [ExcelInvalidValue(-1)]
    public int Age { get; set; }  // -1 if cell value is invalid (e.g., "N/A")
}
```

**Note:** `ExcelInvalidValue` only handles invalid/unparseable values. Empty cells will still throw unless you also use `ExcelDefaultValue` or make the property nullable with `ExcelOptional`.

#### Advanced Fallback Strategies

For more complex fallback scenarios, use `IFallbackItem` types:

```csharp
public class ThrowFallbackItem : IFallbackItem
{
    public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
    {
        throw new InvalidOperationException("Custom error message");
    }
}

public class Employee
{
    public string Name { get; set; }

    // Throw custom exception when cell is empty
    [ExcelEmptyFallback(typeof(ThrowFallbackItem))]
    public int Age { get; set; }

    // Use fallback with constructor arguments when value is invalid
    [ExcelInvalidFallback(typeof(DefaultFallbackItem), ConstructorArguments = new object[] { -1 })]
    public int YearsOfService { get; set; }
}
```

**Available Fallback Attributes:**
- `[ExcelEmptyFallback(Type)]` - Handles empty cells using custom `IFallbackItem`
- `[ExcelInvalidFallback(Type)]` - Handles invalid/unparseable values using custom `IFallbackItem`

Both attributes support `ConstructorArguments` property to pass parameters to the fallback item constructor.

#### Custom Cell Transformers

Transform cell values before mapping using custom transformers:

```csharp
public class UpperCaseTransformer : ICellTransformer
{
    public string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
    {
        return readResult.StringValue?.ToUpperInvariant();
    }
}

public class Employee
{
    [ExcelTransformer(typeof(UpperCaseTransformer))]
    public string Name { get; set; }  // "ALICE JOHNSON"

    // Use built-in trimming transformer
    [ExcelTransformer(typeof(TrimStringCellTransformer))]
    public string Department { get; set; }
}
```

The `ExcelTransformerAttribute` accepts any type implementing `ICellTransformer` and supports `ConstructorArguments` for parameterized transformers.

### Fluent API Mapping

For complex scenarios, use fluent mapping with `ExcelClassMap<T>`:

```csharp
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name)
            .WithColumnName("Full Name");

        Map(e => e.Salary)
            .WithColumnIndex(2);

        Map(e => e.Department)
            .WithColumnNames("Department", "Dept", "Division")
            .MakeOptional();
    }
}

// Register the map
importer.Configuration.RegisterClassMap<EmployeeMap>();

var employees = sheet.ReadRows<Employee>();
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
- `.WithFormats("yyyy-MM-dd", "dd/MM/yyyy")` - Parse dates, times, durations and formats (`DateTime`, `DateTimeOffset`, `TimeSpan`, `DateOnly`, `TimeOnly`) with specific formats
- `.WithMapping(dictionary)` - Map string values to enum/object values
- `.WithElementMap(...)` - Configure element pipeline for collections

#### Complete Fluent Example
```csharp
public enum EmploymentStatus { FullTime, PartTime, Contract }

public class Employee
{
    public string Name { get; set; }
    public EmploymentStatus Status { get; set; }
    public int YearsOfService { get; set; }
    public float PerformanceScore { get; set; }
    public DateTime HireDate { get; set; }
    public string Department { get; set; }
}

public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name);

        // Map misspelled column and string values
        Map(e => e.Status)
            .WithColumnName("Employment Status")
            .WithMapping(new Dictionary<string, EmploymentStatus>
            {
                { "FT", EmploymentStatus.FullTime },
                { "PT", EmploymentStatus.PartTime }
            });

        // Map by index
        Map(e => e.YearsOfService)
            .WithColumnIndex(2);

        // Custom converter
        Map(e => e.PerformanceScore)
            .WithColumnName("Performance (%)")
            .WithConverter(value => float.Parse(value.TrimEnd('%')) / 100f);

        // Date parsing with multiple formats
        Map(e => e.HireDate)
            .WithFormats("yyyy-MM-dd", "dd/MM/yyyy");

        // Try multiple column names
        Map(e => e.Department)
            .WithColumnNames("Dept", "Department", "Division");
    }
}

// Register and use
importer.Configuration.RegisterClassMap<EmployeeMap>();
var employees = sheet.ReadRows<Employee>();
```

## Error Handling

### Nullable Types and Fallbacks

By default:
- Nullable types are set to `null` for empty cells
- Non-nullable types throw `ExcelMappingException` for empty/invalid cells

Configure fallback behavior:

| Name          | Status  | YearsOfService | HireDate    |
|---------------|---------|----------------|-------------|
| Alice Johnson | invalid | invalid        | invalid     |
| Bob Smith     |         |                |             |

```csharp
public enum EmploymentStatus { FullTime, PartTime, Invalid, Unknown }

public class Employee
{
    public string Name { get; set; }
    public EmploymentStatus Status { get; set; }
    public int? YearsOfService { get; set; }
    public DateTime? HireDate { get; set; }
}

public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name);

        Map(e => e.Status)
            .WithEmptyFallback(EmploymentStatus.Unknown)     // Empty cells
            .WithInvalidFallback(EmploymentStatus.Invalid);  // Invalid values

        Map(e => e.YearsOfService)
            .WithInvalidFallback(-1);  // Can't parse as int

        Map(e => e.HireDate)
            .WithInvalidFallback(null);  // Can't parse as DateTime
    }
}

importer.Configuration.RegisterClassMap<EmployeeMap>();
var employees = sheet.ReadRows<Employee>();
```

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

Customize the separator using fluent API:

```csharp
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

#### Fluent Collection Mapping

```csharp
public class Employee
{
    public string Name { get; set; }
    public string[] Skills { get; set; }
    public DateTime[] Certifications { get; set; }
}

public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name);

        // Split by comma (default)
        Map(e => e.Skills)
            .WithColumnName("Skills");

        // Read multiple columns
        Map(e => e.Certifications)
            .WithColumnNames("Certification Date 1", "Certification Date 2")
            .WithElementMap(m => m
                .WithFormats("yyyy-MM-dd", "dd/MM/yyyy")
            );
    }
}
```

#### Supported Collection Types

- `T[]` - Arrays
- `List<T>`, `IList<T>`, `ICollection<T>`, `IEnumerable<T>`
- `HashSet<T>`, `ISet<T>`
- `ImmutableArray<T>`, `ImmutableList<T>`, `ImmutableHashSet<T>`
- `FrozenSet<T>` (.NET 8+)
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

## Special Scenarios

### Sheets Without Headers

Disable header row and use column indices:

|               |            |
|---------------|------------|
| Alice Johnson | Engineering|
| Bob Smith     | Marketing  |

```csharp
public class Employee
{
    public string Name { get; set; }
    public string Department { get; set; }
}

public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name).WithColumnIndex(0);
        Map(e => e.Department).WithColumnIndex(1);
    }
}

using var importer = new ExcelImporter("employees.xlsx");
importer.Configuration.RegisterClassMap<EmployeeMap>();

var sheet = importer.ReadSheet();
sheet.HasHeading = false;  // Disable header row

var employees = sheet.ReadRows<Employee>();
```

### Headers Not in First Row

Skip rows before the header:

|                      |             |
|----------------------|-------------|
| Employee Report 2025 |             |
|                      |             |
| Name                 | Department  |
| Alice Johnson        | Engineering |
| Bob Smith            | Marketing   |

```csharp
public class Employee
{
    public string Name { get; set; }
    public string Department { get; set; }
}

using var importer = new ExcelImporter("employees.xlsx");
var sheet = importer.ReadSheet();
sheet.HeadingIndex = 2;  // Header is on row 3 (zero-based index 2)

var employees = sheet.ReadRows<Employee>();
```

## Performance Tips

1. **Use streaming**: `ReadRows<T>()` uses lazy evaluation - don't materialize unnecessarily
   ```csharp
   // Good - processes one at a time
   foreach (var employee in sheet.ReadRows<Employee>())
   {
       ProcessEmployee(employee);
   }

   // Avoid - loads everything into memory
   var allEmployees = sheet.ReadRows<Employee>().ToList();
   ```

2. **Register maps once**: Class maps are cached per type
   ```csharp
   importer.Configuration.RegisterClassMap<EmployeeMap>();
   ```

3. **Disable blank line skipping**: Off by default for performance
   ```csharp
   importer.Configuration.SkipBlankLines = false;  // Default
   ```

4. **Use column indices for headerless sheets**: Faster than column name lookup
   ```csharp
   Map(e => e.Name).WithColumnIndex(0);  // Faster
   Map(e => e.Name).WithColumnName("Name");  // Requires lookup
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
- `ImmutableArray<T>`, `ImmutableList<T>`
- `Dictionary<TKey, TValue>`, `IDictionary<TKey, TValue>`
- `FrozenDictionary<TKey, TValue>`, `ImmutableDictionary<TKey, TValue>`

### Date and Time Types

ExcelMapper has comprehensive support for modern .NET date and time types:

- `DateTime`, `DateTimeOffset` - Full date and time
- `DateOnly` - Date without time (available in .NET 6+)
- `TimeOnly` - Time without date (available in .NET 6+)
- `TimeSpan` - Duration/time interval

All support custom format parsing with `.WithFormats()`.

### Complex Types
- Classes with public properties/fields
- Record types
- Nested objects
- `ExpandoObject` for dynamic scenarios

## CSV Support

ExcelMapper supports CSV files in addition to Excel formats:

```csharp
// CSV files are automatically detected
using var importer = new ExcelImporter("employees.csv");
var sheet = importer.ReadSheet();
var employees = sheet.ReadRows<Employee>();
```

Supported formats:
- `.xlsx` - Excel 2007+ (Office Open XML)
- `.xls` - Excel 97-2003 (Binary Format)
- `.csv` - Comma-separated values

## Common Issues & Troubleshooting

### Column Not Found

**Problem**: `ExcelMappingException: Could not find column 'ColumnName'`

**Solutions**:
```csharp
// Option 1: Make the property optional
public class Employee
{
    [ExcelOptional]
    public string MiddleName { get; set; }
}

// Option 2: Try multiple column names
public class Employee
{
    [ExcelColumnNames("Department", "Dept", "Division")]
    public string Department { get; set; }
}

// Option 3: Use pattern matching for flexible headers
public class Employee
{
    [ExcelColumnMatching(@"dept.*", RegexOptions.IgnoreCase)]
    public string Department { get; set; }
}
```

### Empty Cell Handling

**Problem**: Exception when reading empty cells into non-nullable types

**Solutions**:
```csharp
// Option 1: Use nullable types
public class Employee
{
    public string Name { get; set; }
    public int? YearsOfService { get; set; }  // Nullable - allows null for empty cells
}

// Option 2: Provide default value
public class Employee
{
    public string Name { get; set; }
    
    [ExcelDefaultValue(0)]
    public int YearsOfService { get; set; }  // Uses 0 for empty cells
}

// Option 3: Use fluent API
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.YearsOfService)
            .WithEmptyFallback(0);
    }
}
```

### Invalid Data Parsing

**Problem**: Exception when cell contains invalid data (e.g., "N/A" in numeric column)

**Solutions**:
```csharp
// Option 1: Use ExcelInvalidValue attribute
public class Employee
{
    [ExcelInvalidValue(-1)]
    public int YearsOfService { get; set; }  // Uses -1 when value can't be parsed
}

// Option 2: Use fluent API for both empty and invalid
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.YearsOfService)
            .WithValueFallback(-1);  // Handles both empty AND invalid
    }
}

// Option 3: Custom converter for complex logic
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.YearsOfService)
            .WithConverter(value => 
            {
                if (string.IsNullOrWhiteSpace(value) || value == "N/A")
                    return -1;
                return int.Parse(value);
            });
    }
}
```

### Case Sensitivity Issues

**Problem**: Column names don't match due to different casing

**Solution**: Column name matching is case-insensitive by default, but you can control enum parsing:

```csharp
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        // Enum parsing can be case-insensitive
        Map(e => e.Status, ignoreCase: true);
    }
}
```

### Date Format Issues

**Problem**: Dates not parsing correctly from different formats

**Solution**: Specify expected date formats:

```csharp
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.HireDate)
            .WithFormats("yyyy-MM-dd", "dd/MM/yyyy", "MM/dd/yyyy");
            
        // Works for DateOnly, TimeOnly, TimeSpan too
        Map(e => e.BirthDate)  // DateOnly property
            .WithFormats("yyyy-MM-dd", "dd/MM/yyyy");
    }
}
```

### Performance Issues with Large Files

**Problem**: Memory or performance issues with large Excel files

**Solutions**:
```csharp
// 1. Use streaming - don't materialize all rows at once
foreach (var employee in sheet.ReadRows<Employee>())
{
    // Process one at a time
    ProcessEmployee(employee);
}

// 2. Disable blank line checking if not needed (it's off by default)
importer.Configuration.SkipBlankLines = false;

// 3. Use column indices instead of names for headerless sheets
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name).WithColumnIndex(0);  // Faster than name lookup
        Map(e => e.Department).WithColumnIndex(1);
    }
}

// 4. Adjust column count limits for very wide sheets
importer.Configuration.MaxColumnsPerSheet = 20000;  // Default is 10000
```

### Multiple Sheets

**Problem**: Need to read specific sheets or multiple sheets

**Solution**:
```csharp
using var importer = new ExcelImporter("workbook.xlsx");

// Read specific sheet by index
var firstSheet = importer.ReadSheet(0);
var employees = firstSheet.ReadRows<Employee>();

// Read specific sheet by name
var engineeringSheet = importer.ReadSheet("Engineering");
var engineers = engineeringSheet.ReadRows<Employee>();

// Read all sheets
foreach (var sheet in importer.ReadSheets())
{
    Console.WriteLine($"Processing sheet: {sheet.Name}");
    var sheetEmployees = sheet.ReadRows<Employee>();
    // Process employees...
}

// Check number of sheets
Console.WriteLine($"Total sheets: {importer.NumberOfSheets}");
```

## Best Practices

1. **Register class maps once** - Class maps are cached, so register them during application startup
   ```csharp
   importer.Configuration.RegisterClassMap<EmployeeMap>();
   ```

2. **Use streaming for large files** - Avoid `.ToList()` or `.ToArray()` unless necessary
   ```csharp
   // Good
   foreach (var employee in sheet.ReadRows<Employee>())
       ProcessEmployee(employee);
   
   // Avoid if dataset is large
   var allEmployees = sheet.ReadRows<Employee>().ToList();
   ```

3. **Handle missing columns gracefully** - Use `[ExcelOptional]` for non-critical columns
   ```csharp
   [ExcelOptional]
   public string MiddleName { get; set; }
   ```

4. **Provide fallback values** - Make your code resilient to data quality issues
   ```csharp
   [ExcelDefaultValue(0)]
   [ExcelInvalidValue(-1)]
   public int YearsOfService { get; set; }
   ```

5. **Use meaningful error messages** - Custom fallback items can provide better diagnostics
   ```csharp
   public class CustomFallback : IFallbackItem
   {
       public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
       {
           throw new InvalidOperationException(
               $"Invalid data in row {rowIndex}, column {readResult.ColumnName}: {readResult.StringValue}"
           );
       }
   }
   ```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

Built on top of [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) for robust Excel file parsing.
