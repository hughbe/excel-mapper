# Mapping

How rows are mapped to objects: automatic mapping, records, attributes, the fluent API and value mapping. See the [documentation index](../README.md#documentation) for other topics.

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

Match columns using regex patterns or custom selectors:

```csharp
public class Employee
{
    public string Name { get; set; }

    // Match columns like "2024 Salary", "2025 Projected Salary"
    [ExcelColumnMatching(@"\d{4}.*Salary", RegexOptions.IgnoreCase)]
    public decimal Salary { get; set; }
}
```

For full control over which columns are read, implement `IColumnSelector`. It's called once for each sheet, and returns the positions of the columns to read (as `ExcelHeading.GetColumnIndex` gives them):

```csharp
public class PrefixSelector : IColumnSelector
{
    private readonly string _prefix;

    public PrefixSelector(string prefix) => _prefix = prefix;

    public IReadOnlyList<int> SelectColumns(ExcelSheet sheet)
        => sheet.Heading!.ColumnNames
            .Where(name => name.StartsWith(_prefix))
            .Select(sheet.Heading.GetColumnIndex)
            .ToArray();

    // Used in error messages, such as: Could not read value for member "Bonuses" for columns starting with "Bonus_".
    public string Describe(ExcelSheet sheet) => $"columns starting with \"{_prefix}\"";
}

public class Employee
{
    // A single value is read from the first column chosen.
    [ExcelColumnMatching(typeof(PrefixSelector), ConstructorArguments = new object[] { "Salary" })]
    public decimal Salary { get; set; }

    // A collection or dictionary is read from every column chosen, in order.
    [ExcelColumnsMatching(typeof(PrefixSelector), ConstructorArguments = new object[] { "Bonus_" })]
    public decimal[] Bonuses { get; set; }
}

// Or with the fluent API
Map(e => e.Bonuses).WithColumns(new PrefixSelector("Bonus_"));
```

- **No columns:** if no column is chosen, the member is missing, as for a column name that isn't found. That throws unless the member is optional.
- **Regions:** when reading a [region](reading.md#reading-part-of-a-sheet), chosen columns outside it are ignored.

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

Or use Excel's column letters, which is the same as the zero-based index (`"A"` is 0, `"C"` is 2):

```csharp
public class Employee
{
    [ExcelColumnLetter("A")]
    public string Name { get; set; }

    [ExcelColumnLetter("B")]
    public int Age { get; set; }
}

// Or with the fluent API
Map(e => e.Name).WithColumnLetter("A");
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

#### Row Numbers

Record which row each object came from, for example to report problems back to whoever filled in the spreadsheet:

```csharp
public class Employee
{
    [ExcelRowNumber]
    public int Row { get; set; }

    public string Name { get; set; }
}

foreach (var employee in sheet.ReadRows<Employee>())
{
    if (string.IsNullOrWhiteSpace(employee.Name))
    {
        Console.WriteLine($"Row {employee.Row} has no name");
    }
}
```

Row numbers are one-based, as Excel shows them, and count every row in the sheet: the heading, any rows above it, and blank rows skipped by `SkipBlankLines`. So the first row under a heading in row 1 is row 2, and the numbers match the rows in error messages. The member doesn't read a column, and can be any numeric type (`int`, `long`, `int?`, ...). With the fluent API, use `MapRowNumber`:

```csharp
importer.Configuration.RegisterClassMap<Employee>(c =>
{
    c.MapRowNumber(e => e.Row);
    c.Map(e => e.Name);
});
```

For a positional record, apply the attribute to the property: `record Employee([property: ExcelRowNumber] int Row, string Name)`.

#### Extra Columns

Keep the columns a class doesn't map, keyed by column name:

| Name  | Age | Team    | Desk |
|-------|-----|---------|------|
| Alice | 34  | Finance | 4B   |

```csharp
public class Employee
{
    public string Name { get; set; }
    public int Age { get; set; }

    [ExcelExtraColumns]
    public Dictionary<string, string> Other { get; set; }  // { ["Team"] = "Finance", ["Desk"] = "4B" }
}
```

A column is extra if no other member of the type being read reads it, including members of nested objects at any depth (an extra columns member on a nested object also leaves out the columns its parent's members read), members with `[ExcelColumnName]`, `[ExcelColumnIndex]` and similar attributes, and split collections. Members with `[ExcelIgnore]` don't read a column, so their columns are extra. Columns with an empty name are skipped. The sheet must have a heading.

The member can be any dictionary with string keys that can be mapped to all columns, such as `Dictionary<string, string>`, `IReadOnlyDictionary<string, object>` (values as Excel stores them: `double`, `DateTime`, ...), `Dictionary<string, int>` or `ExpandoObject`. With the fluent API, use `MapExtraColumns`, which also lets you configure how values are mapped:

```csharp
importer.Configuration.RegisterClassMap<Employee>(c =>
{
    c.Map(e => e.Name).WithColumnName("Full Name");
    c.Map(e => e.Age);
    c.MapExtraColumns(e => e.Other).WithValueMap(v => v.WithConverter(s => s?.Trim()));
});
```

With a class map, only the members in the class map read columns.

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

Cells that Excel stores as dates, times or numbers map directly. For dates and times written as **text** (every CSV cell, or text cells in a workbook), these formats are accepted by default, using the importer's [culture](#culture-for-the-whole-importer) (or the current culture):

| Type | Accepted by default | Examples (en-GB) |
|------|---------------------|------------------|
| `DateTime`, `DateTimeOffset` | ISO 8601; the culture's date and time (`"G"`), short date and time (`"g"`) and short date (`"d"`) | `2024-03-04`, `2024-03-04T10:30:15Z`, `2024-03-04 10:30`, `04/03/2024 10:30:15`, `04/03/2024` |
| `DateOnly` | the culture's short date (`"d"`); ISO 8601 | `04/03/2024`, `2024-03-04` |
| `TimeOnly` | the culture's short time (`"t"`) and long time (`"T"`); 24-hour times | `10:30`, `10:30:15`, `9:05`, `14:30:15.25` |
| `TimeSpan` | the constant format (`"c"`) | `1:30`, `01:30:00`, `1.02:03:04` |

Text such as `3/4/2024 10:30 AM` matches en-US patterns whether it uses an ordinary space or the narrow no-break space that .NET's own formatting uses on Linux, macOS and recent Windows. A `DateTime` with an offset (`Z`, `+01:00`) is converted to local time, as `DateTime.Parse` does; map to `DateTimeOffset` to keep the offset.

Specify formats to accept exactly those formats and nothing else:

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

**Default number styles:**

Text is parsed with `NumberStyles.Number` (leading/trailing whitespace and sign, thousands separators and a decimal point), plus `NumberStyles.AllowExponent` for `float`, `double`, `decimal` and other floating-point types so values such as `1E5` parse. This is the same on every build of the library. Cells that Excel stores as numbers aren't parsed, so number styles don't affect them.

**Common NumberStyles Values:**

- `NumberStyles.Integer` - Simple integers
- `NumberStyles.Float` - Floating-point numbers with exponents
- `NumberStyles.Number` - Includes thousands separator and decimal point
- `NumberStyles.AllowThousands` - Allows thousands separators (e.g., 1,000)
- `NumberStyles.AllowDecimalPoint` - Allows decimal point (e.g., 123.45)
- `NumberStyles.AllowCurrencySymbol` - Allows currency symbols (e.g., $100)
- `NumberStyles.AllowLeadingSign` - Allows leading +/- sign
- `NumberStyles.AllowTrailingSign` - Allows trailing +/- sign
- `NumberStyles.AllowParentheses` - Allows parentheses for negative numbers (e.g., (100))

**Supported Numeric Types:**

The `ExcelNumberStyle` attribute works with these types (and their nullable versions) on every build of the library:
- `int`, `long`, `short`, `byte`, `sbyte`
- `uint`, `ulong`, `ushort`
- `float`, `double`, `decimal`

On .NET 8 and later (the library's `net8.0` build), it also works with `Half`, `Int128`, `UInt128`, `BigInteger`, `nint` and `nuint`, and the fluent `.WithNumberStyle()` and `.WithFormatProvider()` methods are available for numeric members. On .NET Framework and other netstandard runtimes, use the `[ExcelNumberStyle]` attribute, and set the culture with `importer.Configuration.Culture`.

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

To compute the value of an empty or invalid cell, implement `ICellFallback<T>` and set it with `[ExcelEmptyFallback]`/`[ExcelInvalidFallback]` or `.WithEmptyFallback(fallback)`/`.WithInvalidFallback(fallback)`:

```csharp
public class AgeFromBirthYearFallback : ICellFallback<int>
{
    public int Fallback(CellValue value, CellContext context, Exception? exception)
    {
        // Text such as "born 1990" isn't a number: work out the age instead.
        if (value.Text is { } text && text.StartsWith("born ") && int.TryParse(text[5..], out var year))
        {
            return DateTime.Today.Year - year;
        }

        // Throw to report the cell as an error, as the default invalid fallback does.
        throw new ExcelMappingException($"Invalid age \"{value.Text}\".", context.Sheet, context.RowIndex, context.ColumnIndex, exception);
    }
}

public class UnknownYearsFallback(int value) : ICellFallback<int>
{
    public int Fallback(CellValue cell, CellContext context, Exception? exception) => value;
}

public class Employee
{
    public string Name { get; set; }

    [ExcelInvalidFallback(typeof(AgeFromBirthYearFallback))]
    public int Age { get; set; }

    // A fallback for T also works for T? members.
    [ExcelEmptyFallback(typeof(UnknownYearsFallback), ConstructorArguments = new object[] { -1 })]
    public int? YearsOfService { get; set; }
}
```

- **`value`** is the cell's value after any transformers, and **`context`** says where the cell is (see [Custom Converters](#custom-converters)).
- **`exception`** is why the value is invalid, if known. It's null for empty cells.
- **Throwing** reports the cell as an error: `ReadRowResults` records an `ExcelMappingException` rather than throwing it.
- **Both attributes** support `ConstructorArguments`.

#### Custom Converters

To convert cells yourself, implement `ICellConverter<T>` and set it with `[ExcelConverter]` or `.WithConverter(converter)`. The converter replaces the built-in conversion for the member's type:

```csharp
public class ThousandsConverter : ICellConverter<int>
{
    public CellConversionResult<int> Convert(CellValue value, CellContext context)
    {
        // Typed cells keep their type: a number cell is a double, not text.
        if (value.Kind == CellKind.Number)
        {
            return CellConversionResult<int>.Success((int)value.Number);
        }

        var text = value.Text!;
        if (text.EndsWith("k") && int.TryParse(text[..^1], NumberStyles.Integer, context.FormatProvider, out var thousands))
        {
            return CellConversionResult<int>.Success(thousands * 1000);
        }

        // The exception is only created if it's used, for example by the default invalid fallback.
        return CellConversionResult<int>.Invalid(() => new FormatException($"\"{text}\" isn't a number of thousands."));
    }
}

public class Order
{
    [ExcelConverter(typeof(ThousandsConverter))]
    public int Quantity { get; set; }  // "2k" → 2000

    [ExcelConverter(typeof(ThousandsConverter))]
    public int? Reorder { get; set; }  // a converter for T also converts T?

    [ExcelSeparators(';')]
    [ExcelConverter(typeof(ThousandsConverter))]
    public int[] Batches { get; set; }  // converts each element
}

// Or with the fluent API
Map(o => o.Quantity).WithConverter(new ThousandsConverter());
```

- **`CellValue`** is the cell's value: `Kind` (`Empty`, `Text`, `Number`, `Boolean`, `DateTime`, `TimeSpan` or `Other`), `Text`, `Number`, `Boolean`, `DateTime`, `TimeSpan` and `Value`. Numbers, dates and booleans from Excel files arrive typed; CSV cells are always text.
- **`CellContext`** says where the cell is: `Sheet`, `RowIndex`, `ColumnIndex`, `CellAddress` (such as `"B11"`), the `Member` being mapped, and the `FormatProvider` to parse with (the importer's [culture](#culture-for-the-whole-importer), if set).
- **`CellConversionResult<T>`**:
  - `Success(value)`: the converted value.
  - `Invalid()`, `Invalid(exception)` or `Invalid(() => exception)`: the member's invalid fallback is used, which by default throws an `ExcelMappingException` whose inner exception is the one given.
  - `NotHandled`: the next converter is tried. Several `[ExcelConverter]` attributes are tried in order, and if none converts the value, it's invalid.
- **Empty cells** don't reach the converter: they use the member's empty fallback.
- **Converters are shared** between rows and threads, so they mustn't change their own state. `[ExcelConverter]` takes `ConstructorArguments` for converters with parameters.

#### Custom Cell Transformers

Change cell values before they're converted with `ICellTransformer`, set with `[ExcelTransformer]` or `.WithTransformers(...)`:

```csharp
public class UpperCaseTransformer : ICellTransformer
{
    public CellValue Transform(CellValue value, CellContext context)
        => value.Kind == CellKind.Text ? new CellValue(value.Text!.ToUpperInvariant()) : value;
}

// A transformer can also return a typed value: "1 234,5" becomes the number 1234.5.
public class FrenchNumberTransformer : ICellTransformer
{
    private static readonly CultureInfo s_french = CultureInfo.GetCultureInfo("fr-FR");

    public CellValue Transform(CellValue value, CellContext context)
        => value.Kind == CellKind.Text && double.TryParse(value.Text, NumberStyles.Number, s_french, out var number)
            ? new CellValue(number)
            : value;
}

public class Employee
{
    [ExcelTransformer(typeof(UpperCaseTransformer))]
    public string Name { get; set; }  // "ALICE JOHNSON"

    [ExcelTrimString]
    [ExcelTransformer(typeof(FrenchNumberTransformer))]
    public double Salary { get; set; }
}

// Or with the fluent API
Map(e => e.Name).WithTransformers(new UpperCaseTransformer());
```

- **Order:** transformers run in order, before the empty check, so a transformer that returns an empty value (`new CellValue((string?)null)`) makes the cell use the empty fallback.
- **Return values:** return `value` to leave a cell unchanged, or a new `CellValue` holding text, a number, a boolean, a date or a duration.
- **`[ExcelTransformer]`** supports `ConstructorArguments`.

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
- `.WithColumnLetter("C")` - Map to specific column by Excel letter
- `.WithColumnNames("Name1", "Name2")` - Try multiple column names in order
- `.WithColumnIndices(0, 1, 2)` - Try multiple indices in order
- `.WithColumnNameMatching(name => name.Contains("Total"))` - Use predicate
- `.WithColumnNameMatching(new Regex("total", RegexOptions.IgnoreCase))` - Use regular expression
- `.WithColumns(selector)` - Use a custom `IColumnSelector`
- `.WithColumnsMatching(predicate)` or `.WithColumnsMatching(regex)` - Read every matching column into a collection

**Behavior:**
- `.MakeOptional()` - Don't throw if column is missing
- `.WithEmptyFallback(value)` - Use default value if cell is empty
- `.WithInvalidFallback(value)` - Use default value if conversion fails
- `.WithValueFallback(value)` - Use default value for both empty and invalid

**Advanced:**
- `.WithConverter(value => ...)` - Custom conversion delegate. It replaces the built-in conversion for the member's type, so it's used for every cell that isn't empty
- `.WithConverter(converter)` - Convert with an `ICellConverter<T>`, which gets the typed cell value and its context: see [Custom Converters](#custom-converters)
- `.WithFormats("yyyy-MM-dd", "dd/MM/yyyy")` - Parse dates, times, durations and formats (`DateTime`, `DateTimeOffset`, `TimeSpan`, `DateOnly`, `TimeOnly`) with specific formats
- `.WithMapping(dictionary)` - Map string values to enum/object values
- `.WithElementMap(...)` - Configure element pipeline for collections

Every method returns the map, so the calls chain in any order, for example `Map(e => e.HireDate).WithFormats("yyyy-MM-dd").MakeOptional()`.

For nullable members, converters and fallbacks can work with the underlying type: `Map(e => e.Bonus).WithConverter(value => int.Parse(value))` maps an `int?` member, as do an `ICellConverter<int>` and an `ICellFallback<int>`.

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
