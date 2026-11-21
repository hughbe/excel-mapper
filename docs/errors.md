# Errors and Troubleshooting

Error messages, collecting errors, checking columns, fallbacks and common problems. See the [documentation index](../README.md#documentation) for other topics.

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

    Console.WriteLine(ex.CellAddress);  // "C3" (null if the row or column is unknown)
    Console.WriteLine(ex.ColumnName);   // "YearsOfService" (null if unknown)
    Console.WriteLine(ex.Member?.Name); // "YearsOfService" (the property or field, if known)
    Console.WriteLine(ex.RowIndex);     // 2 (zero-based)
    Console.WriteLine(ex.ColumnIndex);  // 2 (zero-based)
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
