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
  - [Value Mapping](#value-mapping)
  - [Collections and Arrays](#collections-and-arrays)
  - [Dictionaries](#dictionaries)
  - [Nested Objects](#nested-objects)
  - [Enums](#enums)
  - [Custom Converters](#custom-converters)
  - [Custom Transformers and Mappers](#custom-transformers-and-mappers)
- [Special Scenarios](#special-scenarios)
  - [Sheets Without Headers](#sheets-without-headers)
  - [Headers Not in First Row](#headers-not-in-first-row)
- [Error Handling](#error-handling)
- [Performance Tips](#performance-tips)
- [Thread Safety](#thread-safety)
- [Supported Types](#supported-types)
- [CSV Support](#csv-support)
- [Common Issues & Troubleshooting](#common-issues--troubleshooting)
- [Best Practices](#best-practices)
- [API Reference](#api-reference)
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

// From existing IExcelDataReader (for advanced scenarios)
using var reader = ExcelReaderFactory.CreateReader(stream);
using var importer = new ExcelImporter(reader);
```

**Advanced: Access the underlying ExcelDataReader**

```csharp
using var importer = new ExcelImporter("data.xlsx");

// Access the underlying reader for advanced scenarios
IExcelDataReader reader = importer.Reader;

// Check number of sheets
int sheetCount = importer.NumberOfSheets;
```

## Reading Sheets

### Read All Sheets

```csharp
foreach (var sheet in importer.ReadSheets())
{
    Console.WriteLine($"Sheet: {sheet.Name}");
    Console.WriteLine($"Visibility: {sheet.Visibility}");  // Visible, Hidden, or VeryHidden
    Console.WriteLine($"Index: {sheet.Index}");
    Console.WriteLine($"Columns: {sheet.NumberOfColumns}");
}
```

**Sheet Visibility:**
- `ExcelSheetVisibility.Visible` - Normal visible sheets
- `ExcelSheetVisibility.Hidden` - Hidden sheets (can be unhidden in Excel)
- `ExcelSheetVisibility.VeryHidden` - Very hidden sheets (requires VBA to unhide)

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
// Read 10 rows starting from row index 5 (after header at index 0)
// Note: startIndex is relative to the beginning of the file, not after the header
var rows = sheet.ReadRows<Employee>(startIndex: 5, count: 10);

// Example: If header is at row 0, data starts at row 1
// startIndex: 1 = first data row
// startIndex: 11 = 11th data row
```

**Important Notes:**
- `startIndex` is the zero-based row index from the start of the sheet
- The `startIndex` must be **after** the header row
- If `HeadingIndex` is 0 (default), `startIndex` must be at least 1
- The method will throw `ExcelMappingException` if rows don't exist

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

ExcelMapper supports three approaches to mapping Excel rows to objects.

### Automatic Mapping

ExcelMapper automatically maps **public properties and fields** by matching column names (case-insensitive by default).

**Important:** 
- Only **public instance properties with setters** are auto-mapped
- Only **public instance fields** are auto-mapped
- Static members, read-only properties, and indexers are ignored
- Use `[ExcelIgnore]` to exclude specific properties/fields
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
    public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
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

#### Value Mapping with Attributes

Map string cell values to specific enum or object values using attributes:

| Name          | Size | Priority |
|---------------|------|----------|
| Alice Johnson | L    | High     |
| Bob Smith     | M    | Med      |

```csharp
public enum TShirtSize { Small, Medium, Large, XLarge }
public enum Priority { Low, Medium, High }

public class Employee
{
    public string Name { get; set; }

    // Map string values to enum using attributes
    [ExcelMappingDictionary("S", TShirtSize.Small)]
    [ExcelMappingDictionary("M", TShirtSize.Medium)]
    [ExcelMappingDictionary("L", TShirtSize.Large)]
    [ExcelMappingDictionary("XL", TShirtSize.XLarge)]
    public TShirtSize Size { get; set; }

    // Handle abbreviations and variations
    [ExcelMappingDictionary("Low", Priority.Low)]
    [ExcelMappingDictionary("L", Priority.Low)]
    [ExcelMappingDictionary("Medium", Priority.Medium)]
    [ExcelMappingDictionary("Med", Priority.Medium)]
    [ExcelMappingDictionary("M", Priority.Medium)]
    [ExcelMappingDictionary("High", Priority.High)]
    [ExcelMappingDictionary("H", Priority.High)]
    public Priority Priority { get; set; }
}

var employees = sheet.ReadRows<Employee>().ToArray();
Console.WriteLine(employees[0].Size);      // Large
Console.WriteLine(employees[1].Priority);  // Medium
```

**Case-Insensitive Matching:**

By default, dictionary key matching is case-sensitive. Use `ExcelMappingDictionaryComparerAttribute` for case-insensitive or custom comparisons:

```csharp
public class Employee
{
    // Case-insensitive matching
    [ExcelMappingDictionary("b", "extra")]
    [ExcelMappingDictionaryComparer(StringComparison.InvariantCultureIgnoreCase)]
    public string Code { get; set; }  // Matches "B", "b" to "extra"
}
```

**Required vs Optional Mapping:**

Control whether unmapped values should cause errors:

```csharp
public class Employee
{
    // Optional (default) - unmapped values pass through as-is
    [ExcelMappingDictionary("FT", "Full Time")]
    [ExcelMappingDictionary("PT", "Part Time")]
    public string Status { get; set; }

    // Required - unmapped values trigger InvalidFallback or throw
    [ExcelMappingDictionary("A", "Active")]
    [ExcelMappingDictionary("I", "Inactive")]
    [ExcelMappingDictionaryBehavior(MappingDictionaryMapperBehavior.Required)]
    [ExcelInvalidValue("Unknown")]
    public string EmploymentStatus { get; set; }
}
```

**Available Behaviors:**
- `MappingDictionaryMapperBehavior.Optional` (default) - Unmapped values pass through unchanged
- `MappingDictionaryMapperBehavior.Required` - Unmapped values are treated as invalid, triggering fallback behavior

### Fluent API Mapping

For complex scenarios, use fluent mapping with `ExcelClassMap<T>`:

**Method 1: Create a class that inherits from ExcelClassMap<T>**

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

**Method 2: Use lambda-based inline configuration**

```csharp
// Configure inline without creating a separate class
importer.Configuration.RegisterClassMap<Employee>(map =>
{
    map.Map(e => e.Name)
        .WithColumnName("Full Name");

    map.Map(e => e.Salary)
        .WithColumnIndex(2);

    map.Map(e => e.Department)
        .WithColumnNames("Department", "Dept", "Division")
        .MakeOptional();
});

var employees = sheet.ReadRows<Employee>();
```

This lambda approach is useful for:
- Quick one-off mappings
- Testing and prototyping
- Dynamic configuration scenarios

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
                { "PT", EmploymentStatus.PartTime },
                { "Contract", EmploymentStatus.Contract },
                { "Contractor", EmploymentStatus.Contract }
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

### Value Mapping

Map string cell values to specific enum or object values using either attributes or the fluent API.

#### Attribute-Based Value Mapping

See [Value Mapping with Attributes](#value-mapping-with-attributes) in the Attribute-Based Mapping section.

#### Fluent API Value Mapping

| Name          | Size | Priority |
|---------------|------|----------|
| Alice Johnson | L    | High     |
| Bob Smith     | M    | Med      |

```csharp
public enum TShirtSize { Small, Medium, Large, XLarge }
public enum Priority { Low, Medium, High }

public class Employee
{
    public string Name { get; set; }
    public TShirtSize Size { get; set; }
    public Priority Priority { get; set; }
}

public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name);
        
        // Map string values to enum
        Map(e => e.Size)
            .WithMapping(new Dictionary<string, TShirtSize>
            {
                { "S", TShirtSize.Small },
                { "M", TShirtSize.Medium },
                { "L", TShirtSize.Large },
                { "XL", TShirtSize.XLarge }
            });
            
        // Handle abbreviations and variations
        Map(e => e.Priority)
            .WithMapping(new Dictionary<string, Priority>
            {
                { "Low", Priority.Low },
                { "L", Priority.Low },
                { "Medium", Priority.Medium },
                { "Med", Priority.Medium },
                { "M", Priority.Medium },
                { "High", Priority.High },
                { "H", Priority.High }
            });
    }
}
```

**Case-Insensitive Matching:**

```csharp
Map(e => e.Code)
    .WithMapping(new Dictionary<string, string>
    {
        { "b", "extra" }
    }, StringComparer.OrdinalIgnoreCase);  // Case-insensitive
```

**Required vs Optional Behavior:**

```csharp
Map(e => e.Status)
    .WithMapping(new Dictionary<string, string>
    {
        { "A", "Active" }
    }, behavior: MappingDictionaryMapperBehavior.Required)  // Unmapped values are invalid
    .WithInvalidFallback("Unknown");
```

This is especially useful for:
- Mapping abbreviations to full values
- Handling data entry variations
- Converting legacy codes to enum values
- Supporting multiple languages or formats

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

## Thread Safety

**Important:** `ExcelSheet` instances maintain mutable state (current row index) and are **not thread-safe**. Each instance should be used by only one thread at a time.

**Safe concurrent processing:**

```csharp
// Read all rows first
var employees = sheet.ReadRows<Employee>().ToList();

// Then process in parallel
Parallel.ForEach(employees, employee =>
{
    ProcessEmployee(employee);
});
```

**Unsafe concurrent processing:**

```csharp
// DON'T DO THIS - Not thread-safe!
Parallel.ForEach(sheet.ReadRows<Employee>(), employee =>
{
    ProcessEmployee(employee);
});
```

**Multiple sheets:**

```csharp
// Each sheet can be processed independently
using var importer = new ExcelImporter("workbook.xlsx");

foreach (var sheet in importer.ReadSheets())
{
    // Each sheet is independent and can be processed separately
    var data = sheet.ReadRows<Employee>().ToList();
    
    // Now safe to process in parallel
    Parallel.ForEach(data, row => ProcessRow(row));
}
```

## Supported Types

### Primitive Types
- **Numeric**: `int`, `long`, `double`, `decimal`, `float`, `byte`, `short`, `uint`, `ulong`, `ushort`, `sbyte`
- **Extended Numeric**: `Int128`, `UInt128`, `BigInteger`, `Half`, `nint` (IntPtr), `nuint` (UIntPtr), `Complex`
- **Text**: `string`, `char`
- **Other**: `bool`, `Guid`, `Uri`, `Version`
- **Enums**: With optional case-insensitive parsing
- **Nullable versions** of all value types
- **Any type implementing `IParsable<T>`** (.NET 7+) - automatically supported
- **Any type implementing `IConvertible`** - automatically supported

### Collection Types
- `T[]`, `List<T>`, `IList<T>`, `ICollection<T>`, `IEnumerable<T>`
- `HashSet<T>`, `ISet<T>`, `FrozenSet<T>`, `ImmutableHashSet<T>`
- `ImmutableArray<T>`, `ImmutableList<T>`
- `Dictionary<TKey, TValue>`, `IDictionary<TKey, TValue>`
- `FrozenDictionary<TKey, TValue>`, `ImmutableDictionary<TKey, TValue>`

### Date and Time Types

ExcelMapper has comprehensive support for modern .NET date and time types:

- `DateTime` - Full date and time with timezone support
- `DateTimeOffset` - Date and time with explicit timezone offset
- `DateOnly` - Date without time (available in .NET 6+)
- `TimeOnly` - Time without date (available in .NET 6+)
- `TimeSpan` - Duration/time interval

All support custom format parsing with `.WithFormats()`.

**Example:**

```csharp
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

### Complex Types
- Classes with public properties/fields
- Record types (C# 9+)
- Nested objects
- `ExpandoObject` for dynamic scenarios

**Record Type Example:**

```csharp
// Records work seamlessly with ExcelMapper
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

## CSV Support

ExcelMapper supports CSV files in addition to Excel formats:

```csharp
// Specify CSV file type explicitly
using var importer = new ExcelImporter("employees.csv", ExcelImporterFileType.Csv);
var sheet = importer.ReadSheet();
var employees = sheet.ReadRows<Employee>();

// Or let ExcelDataReader auto-detect (may work for .csv extension)
using var importer = new ExcelImporter("employees.csv");
```

**Supported file formats:**
- `.xlsx` - Excel 2007+ (Office Open XML)
- `.xls` - Excel 97-2003 (Binary Format)  
- `.xlsb` - Excel Binary Workbook
- `.csv` - Comma-separated values

**Note:** For CSV files, it's recommended to explicitly specify `ExcelImporterFileType.Csv` to ensure proper parsing.

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
       public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
       {
           throw new InvalidOperationException(
               $"Invalid data in row {rowIndex}, column {readResult.ColumnName}: {readResult.StringValue}"
           );
       }
   }
   ```

6. **Dispose resources properly** - Always use `using` statements with `ExcelImporter`
   ```csharp
   using var importer = new ExcelImporter("data.xlsx");
   // Work with importer
   // Automatically disposed at end of scope
   ```

7. **Be mindful of thread safety** - Don't share `ExcelSheet` instances across threads

## API Reference

### Core Classes

**`ExcelImporter`** - Main entry point for reading Excel files
- `ReadSheet()` / `TryReadSheet()` - Read sheets sequentially or by name/index
- `ReadSheets()` - Enumerate all sheets
- `Configuration` - Access configuration for registering maps
- `Reader` - Access underlying `IExcelDataReader` for advanced scenarios
- `NumberOfSheets` - Get total sheet count

**`ExcelSheet`** - Represents a single worksheet
- `ReadRows<T>()` - Read all rows as typed objects (lazy evaluation)
- `ReadRows<T>(startIndex, count)` - Read specific range of rows
- `ReadRow<T>()` / `TryReadRow<T>()` - Read single row
- `ReadHeading()` - Explicitly read header row
- `Name` - Sheet name
- `Visibility` - Sheet visibility (Visible, Hidden, VeryHidden)
- `Index` - Zero-based sheet index
- `NumberOfColumns` - Column count
- `HasHeading` - Whether sheet has header row (default: true)
- `HeadingIndex` - Zero-based index of header row (default: 0)
- `CurrentRowIndex` - Current row being processed

**`ExcelImporterConfiguration`** - Configuration settings
- `RegisterClassMap<T>()` - Register type-specific mapping
- `RegisterClassMap<T>(Action<ExcelClassMap<T>>)` - Inline lambda configuration
- `TryGetClassMap<T>()` - Check if map exists
- `SkipBlankLines` - Skip empty rows (default: false)
- `MaxColumnsPerSheet` - Security limit (default: 10,000)

**`ExcelClassMap<T>`** - Fluent mapping configuration
- `Map(expression)` - Map property or field
- `MapObject<TElement>(expression)` - Map nested object
- `MapEnumerable<TElement>(expression)` - Map collection
- `MapDictionary<TKey, TValue>(expression)` - Map dictionary

### Mapping Attributes

| Attribute | Purpose |
|-----------|---------|
| `[ExcelColumnName("Name")]` | Map to specific column name |
| `[ExcelColumnNames("Name1", "Name2")]` | Try multiple column names |
| `[ExcelColumnIndex(0)]` | Map to column by index |
| `[ExcelColumnIndices(0, 1)]` | Try multiple column indices |
| `[ExcelColumnMatching(@"regex", options)]` | Match columns by pattern |
| `[ExcelColumnsMatching(typeof(Matcher))]` | Custom column matching |
| `[ExcelOptional]` | Don't throw if column missing |
| `[ExcelIgnore]` | Exclude from mapping |
| `[ExcelDefaultValue(value)]` | Default for empty cells |
| `[ExcelInvalidValue(value)]` | Default for invalid values |
| `[ExcelEmptyFallback(typeof(Fallback))]` | Custom empty cell handling |
| `[ExcelInvalidFallback(typeof(Fallback))]` | Custom invalid value handling |
| `[ExcelPreserveFormatting]` | Read formatted string |
| `[ExcelTrimString]` | Auto-trim whitespace |
| `[ExcelTransformer(typeof(Transformer))]` | Apply custom transformer |
| `[ExcelMappingDictionary("key", value)]` | Map string value to enum/object (multiple allowed) |
| `[ExcelMappingDictionaryComparer(comparison)]` | Set string comparison for dictionary keys |
| `[ExcelMappingDictionaryBehavior(behavior)]` | Control required vs optional mapping |

### Fluent API Methods

**Column Selection:**
- `.WithColumnName("Name")` - Map to specific column
- `.WithColumnNames("Name1", "Name2")` - Try multiple names
- `.WithColumnIndex(0)` - Map by index
- `.WithColumnIndices(0, 1, 2)` - Try multiple indices
- `.WithColumnNameMatching(predicate)` - Use predicate
- `.WithColumnMatching(matcher)` - Custom matcher

**Behavior:**
- `.MakeOptional()` - Don't throw if missing
- `.WithEmptyFallback(value)` - Default for empty
- `.WithInvalidFallback(value)` - Default for invalid
- `.WithValueFallback(value)` - Default for both
- `.WithConverter(value => ...)` - Custom conversion
- `.WithFormats("format1", "format2")` - Date/time formats
- `.WithMapping(dictionary)` - Value mapping
- `.WithTrim()` - Trim whitespace
- `.WithTransformers(...)` - Custom transformers
- `.WithMappers(...)` - Custom mappers

**Collections:**
- `.WithSeparators(';', ',')` - Split delimiters (char)
- `.WithSeparators(";", ",")` - Split delimiters (string)
- `.WithElementMap(m => ...)` - Configure element pipeline

### Extension Interfaces

Implement these for advanced customization:

- `ICellMapper` - Custom type conversion logic
- `ICellTransformer` - Transform string values before mapping
- `IFallbackItem` - Custom fallback behavior
- `IExcelColumnMatcher` - Custom column matching logic

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

Built on top of [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) for robust Excel file parsing.
