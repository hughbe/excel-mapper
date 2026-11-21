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
  - [Records and Constructors](#records-and-constructors)
  - [Attribute-Based Mapping](#attribute-based-mapping)
  - [Fluent API Mapping](#fluent-api-mapping)
- [Error Handling](#error-handling)
- [Advanced Features](#advanced-features)
  - [Value Mapping](#value-mapping)
  - [Enums](#enums)
  - [Collections and Arrays](#collections-and-arrays)
  - [Dictionaries](#dictionaries)
  - [Nested Objects](#nested-objects)
  - [Custom Converters](#custom-converters)
  - [Custom Transformers and Mappers](#custom-transformers-and-mappers)
- [Special Scenarios](#special-scenarios)
  - [Sheets Without Headers](#sheets-without-headers)
  - [Headers Not in First Row](#headers-not-in-first-row)
  - [Restricting Data Range](#restricting-data-range)
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

### Code page encodings

ExcelDataReader needs code page encodings such as Windows-1252, which aren't available by default on .NET Core and .NET 5+. `ExcelImporter` registers them automatically (`Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)`), so no setup is needed. Registering them yourself as well is harmless.

### Reader options

Options for the underlying ExcelDataReader reader, such as a password, CSV separators or a fallback encoding, can be passed to the constructor:

```csharp
using ExcelDataReader;

// A password-protected workbook
using var importer = new ExcelImporter("secret.xlsx", ExcelImporterFileType.Excel,
    new ExcelReaderConfiguration { Password = "password" });

// A CSV file that uses semicolons, read as Windows-1252 if it has no byte order mark
using var csvImporter = new ExcelImporter(stream, ExcelImporterFileType.Csv,
    new ExcelReaderConfiguration
    {
        AutodetectSeparators = [';'],
        FallbackEncoding = Encoding.GetEncoding(1252)
    });
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

Like `ReadRows`, these read the heading first if the sheet has one and it hasn't been read yet, so the heading row is never mapped as data.

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

ExcelMapper automatically maps **public properties and fields** by matching column names. Column name matching is **case-insensitive by default** using `StringComparison.OrdinalIgnoreCase`.

**Important:** 
- Only **public instance properties with setters** are auto-mapped, plus read-only properties that match a constructor parameter (see [Records and Constructors](#records-and-constructors))
- Only **public instance fields** are auto-mapped
- Static members, other read-only properties, and indexers are ignored
- Use `[ExcelIgnore]` to exclude specific properties/fields
- Set `importer.Configuration.IgnoreSeparatorsInColumnNames = true` to match `FirstName` to columns such as "First Name", "first_name" or "First-Name" (exact matches are preferred)

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

### Records and Constructors

Types without a public parameterless constructor, such as positional records or classes with get-only properties, are created through a constructor:

| Name          | Department  | Salary |
|---------------|-------------|--------|
| Alice Johnson | Engineering | 95000  |

```csharp
public record Employee(string Name, string Department, decimal Salary);

public class Product
{
    public Product(string name, decimal price)
    {
        Name = name;
        Price = price;
    }

    public string Name { get; }
    public decimal Price { get; }
    public string Label => $"{Name} ({Price:C})";  // Computed, not mapped
}

var employees = sheet.ReadRows<Employee>().ToArray();
```

How the constructor is chosen and called:
- Each parameter is matched to a mapped property or field with the same name (case-insensitive) and a compatible type
- Every parameter must match a mapped member or have a default value; the public constructor matching the most members is used
- Mapped members that aren't constructor parameters are set afterwards, so a record can have extra settable properties
- If an optional column is missing (for example `[property: ExcelOptional]` on a record parameter), the parameter's default value is used
- Read-only properties can be mapped only if they match a constructor parameter
- If the constructor throws, an `ExcelMappingException` is thrown with the original exception as `InnerException`

Column attributes go on the generated property of a positional record:

```csharp
public record Employee(
    [property: ExcelColumnName("Full Name")] string Name,
    [property: ExcelOptional] decimal Salary = 0);
```

Records and get-only classes work with the fluent API as well:

```csharp
importer.Configuration.RegisterClassMap<Product>(map =>
{
    map.Map(p => p.Name).WithColumnName("Product Name");
    map.Map(p => p.Price).WithColumnName("Unit Price");
});
```

### Attribute-Based Mapping

Use attributes to declaratively configure mapping behavior. Column name matching is **case-insensitive by default** (`StringComparison.OrdinalIgnoreCase`).

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

**String Comparison Options:**

Control how column names are matched using `StringComparison`:

```csharp
public class Employee
{
    // Case-insensitive matching (default)
    [ExcelColumnName("Full Name")]
    public string Name { get; set; }

    // Case-sensitive matching
    [ExcelColumnName("Department", StringComparison.Ordinal)]
    public string Department { get; set; }

    // Culture-aware case-insensitive matching
    [ExcelColumnName("Città", StringComparison.CurrentCultureIgnoreCase)]
    public string City { get; set; }
}
```

**Available StringComparison Options:**
- `StringComparison.OrdinalIgnoreCase` (default) - Case-insensitive, culture-invariant
- `StringComparison.Ordinal` - Case-sensitive, culture-invariant
- `StringComparison.CurrentCultureIgnoreCase` - Case-insensitive using current culture
- `StringComparison.CurrentCulture` - Case-sensitive using current culture
- `StringComparison.InvariantCultureIgnoreCase` - Case-insensitive using invariant culture
- `StringComparison.InvariantCulture` - Case-sensitive using invariant culture

#### Multiple Column Name Variants

Try multiple column names in order of preference:

```csharp
public class Employee
{
    public string Name { get; set; }

    // Try these column names in order (case-insensitive by default)
    [ExcelColumnNames("Age", "#Age", "Years")]
    public int Age { get; set; }

    // Or use multiple attributes with different comparison modes
    [ExcelColumnName("Dept", StringComparison.OrdinalIgnoreCase)]
    [ExcelColumnName("Department", StringComparison.Ordinal)]
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

Trimming only affects text. Cells that Excel stores as dates or numbers keep their type, so `[ExcelTrimString]` can be combined with `[ExcelFormats]` or number styles safely.

#### Date and Time Formats

Specify custom formats for parsing date, time, and duration types:

| Event         | EventDate  | StartTime | Duration |
|---------------|------------|-----------|----------|
| Conference    | 2024-03-15 | 09:30     | 02:30:00 |
| Workshop      | 15/03/2024 | 2:00 PM   | 01:15:00 |

```csharp
public class Event
{
    public string Event { get; set; }
    
    [ExcelFormats("yyyy-MM-dd", "dd/MM/yyyy", "MM/dd/yyyy")]
    public DateTime EventDate { get; set; }
    
    [ExcelFormats("HH:mm", "hh:mm tt")]
    public TimeOnly StartTime { get; set; }
    
    [ExcelFormats(@"hh\:mm\:ss", @"mm\:ss")]
    public TimeSpan Duration { get; set; }
}
```

Or use the fluent API:

```csharp
public class EventMap : ExcelClassMap<Event>
{
    public EventMap()
    {
        Map(e => e.EventDate)
            .WithFormats("yyyy-MM-dd", "dd/MM/yyyy", "MM/dd/yyyy");
            
        Map(e => e.StartTime)
            .WithFormats("HH:mm", "hh:mm tt");
            
        Map(e => e.Duration)
            .WithFormats(@"hh\:mm\:ss", @"mm\:ss");
    }
}
```

**Supported Types:**
- `DateTime` / `DateTime?`
- `DateTimeOffset` / `DateTimeOffset?`
- `DateOnly` / `DateOnly?` (.NET 6+)
- `TimeOnly` / `TimeOnly?` (.NET 6+)
- `TimeSpan` / `TimeSpan?`

#### Culture for the Whole Importer

By default, text is parsed and numbers and dates are converted to text using the current culture of the thread reading the rows, so results can depend on the machine's settings. Set a culture for the whole importer to make them predictable:

```csharp
using var importer = new ExcelImporter("data.xlsx");
importer.Configuration.Culture = CultureInfo.InvariantCulture;   // or new CultureInfo("de-DE")
```

The culture is only used while each row is mapped; your own code between rows keeps its culture. A format provider set on an individual member (see below) takes precedence.

#### Culture-Specific Date and Time Parsing

Specify a format provider (culture) for parsing date, time, and duration types. This is particularly useful when your Excel data uses locale-specific formatting:

| Event         | EventDate  | StartTime | 
|---------------|------------|-----------|
| Conference    | 15.03.2024 | 14:30     |
| Workshop      | 22.11.2024 | 09:15     |

```csharp
using System.Globalization;

public class Event
{
    public string Event { get; set; }
    public DateTime EventDate { get; set; }
    public TimeOnly StartTime { get; set; }
}

public class EventMap : ExcelClassMap<Event>
{
    public EventMap()
    {
        var germanCulture = new CultureInfo("de-DE");
        
        Map(e => e.Event);
        
        Map(e => e.EventDate)
            .WithFormats("dd.MM.yyyy")
            .WithFormatProvider(germanCulture);
            
        Map(e => e.StartTime)
            .WithFormats("HH:mm")
            .WithFormatProvider(germanCulture);
    }
}

importer.Configuration.RegisterClassMap<EventMap>();
var events = sheet.ReadRows<Event>();
```

**Why use format providers?**
- Handle locale-specific number formats (1.234,56 vs 1,234.56)
- Parse dates with culture-specific month/day names
- Support culture-specific date separators (. vs / vs -)
- Ensure correct parsing of time formats across cultures

**Supported Types with Format Providers:**
- `DateTime` / `DateTime?`
- `DateTimeOffset` / `DateTimeOffset?`
- `DateOnly` / `DateOnly?` (.NET 6+)
- `TimeOnly` / `TimeOnly?` (.NET 6+)
- `TimeSpan` / `TimeSpan?`

#### Number Parsing with Number Styles

Control how numeric values are parsed using `NumberStyles` to handle thousands separators, currency symbols, and other formatting. This is particularly useful for financial data or numbers formatted with locale-specific conventions:

| ProductName | Price     | Quantity |
|-------------|-----------|----------|
| Widget A    | 1,234.56  | 1000     |
| Widget B    | $2,500.00 | 500      |

```csharp
using System.Globalization;

public class Product
{
    public string ProductName { get; set; }
    
    [ExcelNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)]
    [ExcelDefaultValue(0.0)]
    [ExcelInvalidValue(0.0)]
    public decimal Price { get; set; }
    
    [ExcelNumberStyle(NumberStyles.AllowThousands)]
    [ExcelDefaultValue(0)]
    [ExcelInvalidValue(0)]
    public int Quantity { get; set; }
}

var products = sheet.ReadRows<Product>().ToArray();
// Successfully parses "1,234.56" as 1234.56 and "1000" as 1000
```

**Using Fluent API:**

For more complex scenarios, use the fluent API with `.WithNumberStyle()` and `.WithFormatProvider()`:

```csharp
public class ProductMap : ExcelClassMap<Product>
{
    public ProductMap()
    {
        Map(p => p.ProductName);
        
        Map(p => p.Price)
            .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
            .WithEmptyFallback(0.0m)
            .WithInvalidFallback(0.0m);
            
        Map(p => p.Quantity)
            .WithNumberStyle(NumberStyles.AllowThousands)
            .WithEmptyFallback(0)
            .WithInvalidFallback(0);
    }
}

importer.Configuration.RegisterClassMap<ProductMap>();
```

**Combining Number Styles with Format Providers:**

When working with locale-specific number formats (e.g., European format using comma as decimal separator):

| Price (EUR) |
|-------------|
| 1.234,56    |
| 2.500,00    |

```csharp
public class ProductMap : ExcelClassMap<Product>
{
    public ProductMap()
    {
        var europeanFormat = new NumberFormatInfo
        {
            NumberGroupSeparator = ".",  // Period for thousands
            NumberDecimalSeparator = "," // Comma for decimals
        };
        
        Map(p => p.Price)
            .WithNumberStyle(NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint)
            .WithFormatProvider(europeanFormat)
            .WithEmptyFallback(0.0m)
            .WithInvalidFallback(0.0m);
    }
}
```

**Common NumberStyles Values:**

- `NumberStyles.Integer` - Simple integers (default for integer types)
- `NumberStyles.Float` - Floating-point numbers (default for float/double)
- `NumberStyles.Number` - Includes thousands separator and decimal point
- `NumberStyles.AllowThousands` - Allows thousands separators (e.g., 1,000)
- `NumberStyles.AllowDecimalPoint` - Allows decimal point (e.g., 123.45)
- `NumberStyles.AllowCurrencySymbol` - Allows currency symbols (e.g., $100)
- `NumberStyles.AllowLeadingSign` - Allows leading +/- sign
- `NumberStyles.AllowTrailingSign` - Allows trailing +/- sign
- `NumberStyles.AllowParentheses` - Allows parentheses for negative numbers (e.g., (100))

**Supported Numeric Types:**

The `ExcelNumberStyle` attribute and `WithNumberStyle()` method work with:
- `int`, `long`, `short`, `byte`, `sbyte`
- `uint`, `ulong`, `ushort`
- `float`, `double`, `decimal`
- `Half` (.NET 5+)
- `Int128`, `UInt128` (.NET 7+)
- `BigInteger`
- `IntPtr` (`nint`), `UIntPtr` (`nuint`)
- Nullable versions of all above types

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

#### Data Annotations Validation

Validate mapped objects using standard .NET data annotations:

| Name          | Percentage |
|---------------|------------|
| Alice Johnson | 85         |
| Bob Smith     | 150        |

```csharp
using System.ComponentModel.DataAnnotations;

public class Employee
{
    public string Name { get; set; }

    [Range(1, 100)]
    public int Percentage { get; set; }
}

// Enable data annotations validation
using var importer = new ExcelImporter("employees.xlsx");
importer.Configuration.ValidateDataAnnotations = true;

var sheet = importer.ReadSheet();
var row1 = sheet.ReadRow<Employee>();  // Success: 85 is valid
var row2 = sheet.ReadRow<Employee>();  // Throws: 150 is out of range
```

**Supported Data Annotations:**
- `[Required]` - Property must have a value
- `[Range(min, max)]` - Value must be within range
- `[StringLength(max, MinimumLength = min)]` - String length constraints
- `[RegularExpression(pattern)]` - Value must match pattern
- `[EmailAddress]`, `[Phone]`, `[Url]` - Format validators
- `[Compare(other)]` - Compare with another property
- `[CreditCard]` - Credit card number validation
- Any custom `ValidationAttribute`

**Note:** Validation occurs **after** mapping is complete and **before** the object is returned. When validation fails, a `System.ComponentModel.DataAnnotations.ValidationException` is thrown with details about the validation error.

**Error Handling:**

```csharp
using System.ComponentModel.DataAnnotations;

try
{
    var employee = sheet.ReadRow<Employee>();
}
catch (ValidationException ex)
{
    // ex.Message contains the validation error message:
    // "The field Percentage must be between 1 and 100."
    Console.WriteLine($"Validation failed: {ex.Message}");
}
```

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
- `.WithColumnNameMatching(new Regex("total", RegexOptions.IgnoreCase))` - Use regular expression
- `.WithColumnMatching(matcher)` - Use custom `IExcelColumnMatcher`
- `.WithColumnsMatching(predicate)`, `.WithColumnsMatching(regex)` or `.WithColumnsMatching(matcher)` - Read every matching column into a collection

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

### Error Messages and Cell Locations

When a cell can't be mapped, `ExcelMappingException` reports where the problem is the way Excel shows it: a cell address when the column is known, otherwise a one-based row number.

```csharp
try
{
    var employees = sheet.ReadRows<Employee>().ToArray();
}
catch (ExcelMappingException ex)
{
    // "Cannot assign "abc" to member "YearsOfService" of type "System.Int32"
    //  in column "YearsOfService" at cell C3 in sheet "Employees"."
    Console.WriteLine(ex.Message);

    Console.WriteLine(ex.CellAddress); // "C3" (null if the row or column is unknown)
    Console.WriteLine(ex.RowIndex);    // 2 (zero-based)
    Console.WriteLine(ex.ColumnIndex); // 2 (zero-based)
}
```

### Checking Columns Before Reading

A missing column is otherwise reported when the first row is read, one column at a time. To check every column a type needs before reading any rows, validate the heading:

```csharp
var sheet = importer.ReadSheet();

// Throws an ExcelMappingException listing every missing column, e.g.
// Sheet "Employees" is missing columns needed to map type "Employee":
//   column "Salary" (member Salary); column "Zip" (member Address.Zip).
sheet.ValidateHeading<Employee>();

// Or get the list without an exception
if (!sheet.TryValidateHeading<Employee>(out var problems))
{
    foreach (var problem in problems)
    {
        Console.WriteLine(problem);   // column "Salary" (member Salary)
    }
}
```

Optional members (`[ExcelOptional]`, `.MakeOptional()`) aren't required. Nested objects, collections and indexed members are checked too, and if a column's name differs only by spaces, underscores, hyphens or dots, the problem suggests it, such as `column "Salary", did you mean "Sal-ary"?`. Validating reads the heading but no rows.

### Collecting Errors

`ReadRows` throws at the first cell that can't be mapped. To check a whole file and report every problem, for example in an import screen, use `ReadRowResults`. It records every invalid cell in every row and keeps going:

```csharp
var results = sheet.ReadRowResults<Employee>().ToList();

var employees = results.Where(r => r.IsValid).Select(r => r.Value!).ToList();

foreach (var result in results.Where(r => !r.IsValid))
{
    foreach (var error in result.Errors)
    {
        // e.g. "C5: Cannot assign "abc" to member "Salary" of type "System.Decimal" ..."
        Console.WriteLine($"{error.CellAddress}: {error.Message}");
    }
}
```

- `Value` is the mapped object, or `default` if the row has errors (partially mapped objects are never returned)
- `Errors` lists every problem in the row: invalid or missing values, exceptions thrown by a constructor, and failed data annotations when `ValidateDataAnnotations` is enabled
- `RowIndex` is the zero-based row in the sheet; each error's `CellAddress` gives the cell, such as `"C5"`, where known
- Problems with the mapping itself, such as a type that can't be auto-mapped, still throw immediately
- A missing column is reported on every row, so call `ValidateHeading` first to report missing columns once

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

Types that don't look up columns by name can be auto-mapped without a heading. Primitives are read from the first column, and members can use `[ExcelColumnIndex]`:

```csharp
public record Employee(
    [property: ExcelColumnIndex(0)] string Name,
    [property: ExcelColumnIndex(1)] string Department);

sheet.HasHeading = false;
var employees = sheet.ReadRows<Employee>();  // No class map needed
var ids = otherSheet.ReadRows<int>();        // Reads the first column
```

Auto-mapping a type whose members are matched by column name throws an `ExcelMappingException` on a sheet without a heading.

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

### Restricting Data Range

Limit which rows and columns are processed using `DataRange`:

|                      |             |            |
|----------------------|-------------|------------|
| Employee Report 2025 |             |            |
| Name                 | Department  | Obsolete   |
| Alice Johnson        | Engineering | Old Data   |
| Bob Smith            | Marketing   | Old Data   |
| Carol White          | Sales       | Old Data   |
|                      |             |            |
| Footer: Generated... |             |            |

```csharp
using var importer = new ExcelImporter("employees.xlsx");
var sheet = importer.ReadSheet();

// Specify exact row and column ranges (zero-based indices)
sheet.DataRange = new ExcelRange(
    rowStart: 1,      // Start at row 2 (after row 0)
    rowEnd: 4,        // End at row 5 (inclusive)
    columnStart: 0,   // Start at column A
    columnEnd: 1      // End at column B (inclusive)
);

var employees = sheet.ReadRows<Employee>();
// Will read only rows 2-5, columns A-B
// Skips header row 0, footer row 6, and column C
```

**Using Excel-style address notation:**

```csharp
// Parse Excel address strings like "A1:B10"
sheet.DataRange = new ExcelRange("B2:D10");  // Columns B-D, rows 2-10

// Single cell
sheet.DataRange = new ExcelRange("A1");  // Just cell A1

// Column ranges
sheet.DataRange = new ExcelRange("A:C");  // All rows, columns A-C

// Row ranges  
sheet.DataRange = new ExcelRange("5:20");  // Rows 5-20, all columns

// Absolute references are accepted too
sheet.DataRange = new ExcelRange("$B$2:$D$10");  // Same as "B2:D10"
```

Column letters are case-insensitive. Address strings must stay within the limits of an Excel worksheet (rows 1 to 1,048,576 and columns A to XFD). For larger CSV files, use the `Range` or index-based constructors instead.

**Important Notes:**
- `DataRange` must be set **before** calling `ReadHeading()` or `ReadRows()`
- Row indices are zero-based (row 0 = first row)
- When `HasHeading` is true, the heading row is at `DataRange.Rows.Start`
- Data rows start **after** the heading row
- Both row and column ranges are inclusive (start and end indices are included)
- Use `Range.All` for unbounded ranges (e.g., `new ExcelRange(Range.All, 0..5)`)

**Common Use Cases:**

```csharp
// Skip header and footer rows
sheet.DataRange = new ExcelRange("A2:Z100");

// Process only specific columns
sheet.DataRange = new ExcelRange("B:E");  // Columns B, C, D, E

// Limit data to first 1000 rows
sheet.DataRange = new ExcelRange(
    rowStart: 0,
    rowEnd: 999,
    columnStart: 0,
    columnEnd: null  // No column limit
);

// Complex scenario: Header in row 3, data in rows 4-50, columns C-F
var sheet = importer.ReadSheet();
sheet.HeadingIndex = 3;  // Header at row 3
sheet.DataRange = new ExcelRange("C3:F50");  // Range includes header and data
```

**Parse and TryParse methods:**

```csharp
// Parse - throws on invalid format
var range1 = ExcelRange.Parse("A1:B10");

// TryParse - returns false on invalid format
if (ExcelRange.TryParse("A1:B10", out var range2))
{
    sheet.DataRange = range2;
}

// ToString returns the Excel address when the range can be written as one
Console.WriteLine(ExcelRange.Parse("$B$2:$D$10")); // "B2:D10"

// Ranges compare by value
Console.WriteLine(ExcelRange.Parse("A:C") == new ExcelRange(Range.All, 0..3)); // True

// default(ExcelRange) means every row and column, the same as new ExcelRange()
sheet.DataRange = default;
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

**Sharing class maps between importers:**

A class map instance can be registered in many importers, including importers used on different threads at the same time. Maps don't keep sheets alive, so a single shared map is safe in a long-running service:

```csharp
private static readonly EmployeeMap s_employeeMap = new();

public IReadOnlyList<Employee> Import(Stream stream)
{
    using var importer = new ExcelImporter(stream);
    importer.Configuration.RegisterClassMap(s_employeeMap);
    return importer.ReadSheet().ReadRows<Employee>().ToList();
}
```

The importer itself, and each sheet, should still be used by one thread at a time.

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

> **Note:** Fractional numeric cells (such as `2.7`) mapped to integer types are invalid rather than truncated or rounded. See [Fractional Numbers Mapped to Integer Properties](#fractional-numbers-mapped-to-integer-properties).

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
- `DateTimeOffset` - Date and time with an offset (cells without one use the machine's local offset)
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

## CSV Support

ExcelMapper supports CSV files in addition to Excel formats:

```csharp
// CSV files must be opened with ExcelImporterFileType.Csv
using var importer = new ExcelImporter("employees.csv", ExcelImporterFileType.Csv);
var sheet = importer.ReadSheet();
var employees = sheet.ReadRows<Employee>();
```

Comma, semicolon, tab, pipe and `#` separators are detected automatically. To use a specific separator or encoding, pass an `ExcelReaderConfiguration` (see [Reader options](#reader-options)).

**Supported file formats:**
- `.xlsx` - Excel 2007+ (Office Open XML)
- `.xls` - Excel 97-2003 (Binary Format)  
- `.xlsb` - Excel Binary Workbook
- `.csv` - Comma-separated values

**Note:** `new ExcelImporter(path)` only detects Excel formats. Opening a CSV file without `ExcelImporterFileType.Csv` throws a `HeaderException`.

## Common Issues & Troubleshooting

### "No data is available for encoding 1252"

**Problem**: `NotSupportedException: No data is available for encoding 1252` is thrown when creating an ExcelDataReader reader.

**Solution**: `ExcelImporter` registers code page encodings automatically. If you create an `IExcelDataReader` yourself (for example to pass to `new ExcelImporter(reader)`) before creating any `ExcelImporter`, register them first:

```csharp
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
```

### Column Not Found

**Problem**: `ExcelMappingException: Could not read value for member "FirstName" for column "FirstName" on row 2 in sheet "Sheet1".`

If a column's name only differs by spaces, underscores, hyphens or dots, the message suggests it, for example `Did you mean column "First Name"?`.

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

// Option 4: Ignore spaces, underscores, hyphens and dots in column names,
// so FirstName matches "First Name", "first_name" or "First-Name"
importer.Configuration.IgnoreSeparatorsInColumnNames = true;
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

### Fractional Numbers Mapped to Integer Properties

**Problem**: A numeric cell such as `2.7` mapped to an `int`, `long`, `byte` etc. property throws an `ExcelMappingException`.

Excel stores every number as a floating-point value, so a cell holding `2.7` is read as the `double` `2.7`. ExcelMapper doesn't silently drop or round the fractional part: the cell is invalid, just as the text `"2.7"` can't be parsed as an `int`. Whole numbers such as `2.0` map normally. This is the same on every runtime.

> **Changed in 4.0:** earlier versions truncated (`2.7` → `2`) on .NET 8 and rounded (`2.7` → `3`) on older runtimes, without an error.

**Solutions**:
```csharp
// Option 1: Map to a floating-point or decimal property
public decimal Quantity { get; set; }

// Option 2: Use a fallback value for fractional (and other invalid) values
[ExcelInvalidValue(0)]
public int Quantity { get; set; }

// Option 3: Take control of the conversion with a custom converter
Map(e => e.Quantity)
    .WithConverter(value => (int)Math.Round(double.Parse(value), MidpointRounding.AwayFromZero));
```

### Case Sensitivity Issues

**Problem**: Column names don't match due to different casing or culture-specific characters

**Solution**: Column name matching is case-insensitive by default (`StringComparison.OrdinalIgnoreCase`), but you can customize it:

```csharp
// Using attributes - specify comparison mode
public class Employee
{
    // Case-sensitive matching
    [ExcelColumnName("Name", StringComparison.Ordinal)]
    public string Name { get; set; }

    // Culture-aware matching (for non-English characters)
    [ExcelColumnName("Città", StringComparison.CurrentCultureIgnoreCase)]
    public string City { get; set; }
}

// Enum parsing can also be case-insensitive
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
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
           throw new ExcelMappingException(
               $"Invalid data \"{readResult.StringValue}\" for {member?.Name}",
               sheet, rowIndex, readResult.ColumnIndex, exception);
           // Message: Invalid data "abc" for Salary in column "Salary" at cell D5 in sheet "Employees".
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
- `ReadRowResults<T>()` - Read all rows, reporting every invalid cell instead of throwing
- `ReadHeading()` - Explicitly read header row
- `ValidateHeading<T>()` / `TryValidateHeading<T>()` - Check that every column a type needs is present, before reading rows
- `Name` - Sheet name
- `Visibility` - Sheet visibility (Visible, Hidden, VeryHidden)
- `Index` - Zero-based sheet index
- `NumberOfColumns` - Column count
- `HasHeading` - Whether sheet has header row (default: true)
- `HeadingIndex` - Zero-based index of header row (default: 0)
- `DataRange` - Range of rows and columns to process (default: all)
- `CurrentRowIndex` - Current row being processed

**`ExcelRange`** - Represents a range of rows and columns
- Constructor: `new ExcelRange(rowStart, rowEnd, columnStart, columnEnd)` - Create range with explicit indices
- Constructor: `new ExcelRange(rows, columns)` - Create range with Range objects
- Constructor: `new ExcelRange(address)` - Create range from Excel address string (e.g., "A1:B10")
- `Parse(string)` - Parse Excel address string (throws on error)
- `TryParse(string, out ExcelRange)` - Try parse Excel address string (returns bool)
- `Rows` - Row range (System.Range)
- `Columns` - Column range (System.Range)
- `ToString()` - Excel address such as "A1:B10", "A:C" or "1:5" when the range can be written as one
- Equality: `Equals`, `==` and `!=` compare rows and columns; `default(ExcelRange)` means all rows and columns

**`ExcelMappingException`** - Thrown when a sheet or cell can't be mapped
- `CellAddress` - Excel address of the cell, such as "B11" (null if unknown)
- `RowIndex` / `ColumnIndex` - Zero-based row and column of the cell (-1 if unknown)
- `Sheet` - The sheet being read

**`ExcelImporterConfiguration`** - Configuration settings
- `RegisterClassMap<T>()` - Register type-specific mapping. Replaces a map created earlier by auto-mapping; throws if you already registered a map for the type
- `RegisterClassMap<T>(Action<ExcelClassMap<T>>)` - Inline lambda configuration
- `TryGetClassMap<T>()` - Check if map exists
- `SkipBlankLines` - Skip empty rows (default: false)
- `IgnoreSeparatorsInColumnNames` - Ignore spaces, underscores, hyphens and dots when matching column names (default: false)
- `Culture` - Culture used to parse text and convert numbers and dates to text (default: null, the current culture)
- `MaxColumnsPerSheet` - Security limit (default: 10,000)
- `ValidateDataAnnotations` - Enable data annotations validation (default: false)

**`ExcelClassMap<T>`** - Fluent mapping configuration
- `Map(expression)` - Map property or field
- `MapObject<TElement>(expression)` - Map nested object
- `MapEnumerable<TElement>(expression)` - Map collection
- `MapDictionary<TKey, TValue>(expression)` - Map dictionary

### Mapping Attributes

| Attribute | Purpose |
|-----------|---------|
| `[ExcelColumnName("Name")]` | Map to specific column name (case-insensitive by default) |
| `[ExcelColumnName("Name", StringComparison.Ordinal)]` | Map to column name with specific comparison mode |
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
| `[ExcelFormats("format1", "format2")]` | Parse dates/times with specific formats |
| `[ExcelNumberStyle(NumberStyles.AllowThousands)]` | Control numeric parsing (thousands, currency, etc.) |
| `[ExcelUri(UriKind.RelativeOrAbsolute)]` | Specify URI kind (Absolute, Relative, or RelativeOrAbsolute) |
| `[ExcelSeparators(';', ',')]` | Split cell value with custom separators |
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
- `.WithColumnNameMatching(predicate)` / `.WithColumnNameMatching(regex)` - Use predicate or regular expression
- `.WithColumnMatching(matcher)` - Custom matcher
- `.WithColumnsMatching(predicate | regex | matcher)` - All matching columns (collections and dictionaries)

**Behavior:**
- `.MakeOptional()` - Don't throw if missing
- `.WithEmptyFallback(value)` - Default for empty
- `.WithInvalidFallback(value)` - Default for invalid
- `.WithValueFallback(value)` - Default for both
- `.WithConverter(value => ...)` - Custom conversion
- `.WithFormats("format1", "format2")` - Date/time formats
- `.WithFormatProvider(IFormatProvider)` - Culture-specific parsing for dates/times/numbers
- `.WithNumberStyle(NumberStyles)` - Control numeric parsing (thousands, currency, etc.)
- `.WithUriKind(UriKind)` - Specify URI kind for Uri properties
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
