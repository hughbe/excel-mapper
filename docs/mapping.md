# Mapping

How rows become objects: mapping by convention, records, attributes and class maps; choosing columns; converting values yourself; and what happens to empty and invalid cells. How each type is parsed (numbers, dates, enums, collections, ...) is in [Types](types.md). See the [documentation index](../README.md#documentation) for other topics.

## Three ways to map

**By convention.** With no configuration, each public property and field is read from the column with the same name, ignoring case:

| Name          | Department  | HireDate   | Salary |
|---------------|-------------|------------|--------|
| Alice Johnson | Engineering | 2020-03-15 | 95000  |
| Bob Smith     | Marketing   | 2019-07-22 | 78000  |

```csharp
public class Employee
{
    public string Name { get; set; }
    public string Department { get; set; }
    public DateTime HireDate { get; set; }
    public decimal Salary { get; set; }
}

var employees = sheet.ReadRows<Employee>().ToArray();
Console.WriteLine(employees[1].Salary); // 78000
```

Public instance properties with a setter, and public instance fields, are mapped. Read-only properties are mapped only through a constructor (see [Records and constructors](#records-and-constructors)); static members and indexers are ignored. Set `importer.Configuration.IgnoreSeparatorsInColumnNames = true` to also match `FirstName` to columns such as "First Name", "first_name" or "First-Name" (an exact match is preferred).

If more than one column matches a member's name (two columns named "Name", or "Name" and "name"), reading it throws an `ExcelMappingException` naming both columns, as it isn't clear which to read. Choose one by its letter or index (`[ExcelColumnLetter("B")]`, `.WithColumnLetter("B")`), or rename one of the columns. Columns with the same name that no member reads, such as several blank heading cells, are fine.

**With attributes** on the type, to change how a member is read:

```csharp
public class Employee
{
    [ExcelColumnName("Full Name")]
    public string Name { get; set; }

    [ExcelOptional]
    public string Department { get; set; }

    [ExcelFormats("dd/MM/yyyy")]
    public DateTime HireDate { get; set; }

    [ExcelInvalidValue(0)]
    public decimal Salary { get; set; }
}
```

**With a class map**, to keep the type free of attributes or configure it in code. Derive from `ExcelClassMap<T>` and map each member in the constructor:

```csharp
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name).WithColumnName("Full Name");
        Map(e => e.Department).MakeOptional();
        Map(e => e.HireDate).WithFormats("dd/MM/yyyy");
        Map(e => e.Salary).WithInvalidFallback(0m);
    }
}

importer.Configuration.RegisterClassMap<EmployeeMap>();
```

Or configure one inline:

```csharp
importer.Configuration.RegisterClassMap<Employee>(c =>
{
    c.Map(e => e.Name).WithColumnName("Full Name");
});
```

A class map reads only the members it maps, and attributes on those members still apply. To start from what auto-mapping would map and change only some members, call `AutoMap()` first: members mapped after it replace the auto-mapped ones.

```csharp
importer.Configuration.RegisterClassMap<Employee>(c =>
{
    c.AutoMap();
    c.Map(e => e.Name).WithColumnName("Full Name");
});
```

Each fluent method returns the map, so calls chain in any order. Every method and attribute is listed in the [API reference](api.md).

Configure a class map before reading with it. The first time rows are read with a class map (or `ValidateHeading` or `FindHeading` uses it), it's frozen: changing it, or the map of one of its members, throws an `InvalidOperationException`. This is what lets one class map be shared between importers and threads.

## Records and constructors

Types without a public parameterless constructor, such as positional records and classes with get-only properties, are created through a constructor:

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
    public string Label => $"{Name} ({Price:C})"; // Not a constructor parameter, so not mapped
}
```

- Each parameter is matched to a mapped member with the same name (ignoring case) and a compatible type. Every parameter must match a member or have a default value, and the public constructor that matches the most members is used.
- Mapped members that aren't parameters are set after the object is created, so a record can have extra settable properties.
- If an optional column is missing, the parameter's default value is used.
- If the constructor throws, the `ExcelMappingException` for the row has its exception as `InnerException`.

Put attributes for a positional record's members on the generated property:

```csharp
public record Employee(
    [property: ExcelColumnName("Full Name")] string Name,
    [property: ExcelOptional] decimal Salary = 0);
```

Class maps work the same way: `c.Map(p => p.Name)` maps the `Name` parameter.

## Choosing columns

| To read | Attribute | Fluent method |
|---|---|---|
| A column by name (ignoring case) | `[ExcelColumnName("Full Name")]` | `.WithColumnName("Full Name")` |
| A column by name, compared another way | `[ExcelColumnName("Name", StringComparison.Ordinal)]` | `.WithColumnName("Name", StringComparison.Ordinal)` |
| The first of several possible columns | `[ExcelColumnNames("Dept", "Department")]` | `.WithColumnNames("Dept", "Department")` |
| A column whose name matches a pattern | `[ExcelColumnMatching(@"\d{4} Salary")]` | `.WithColumnNameMatching(regex)` or `(name => ...)` |
| A column by position | `[ExcelColumnIndex(0)]` or `[ExcelColumnLetter("A")]` | `.WithColumnIndex(0)` or `.WithColumnLetter("A")` |
| The first of several positions | `[ExcelColumnIndices(2, 1)]` | |
| Columns chosen in code | `[ExcelColumnMatching(typeof(Selector))]` | `.WithColumns(selector)` |

Collections and dictionaries read several columns, or split one cell: see [Collections](types.md#collections) and [Dictionaries](types.md#dictionaries).

Positions are zero-based and count from the first column of the sheet (or of the [region](reading.md#reading-part-of-a-sheet) being read). They work without a heading:

```csharp
public class Employee
{
    [ExcelColumnLetter("A")]
    public string Name { get; set; }

    [ExcelColumnIndex(1)]
    public int Age { get; set; }
}

sheet.HasHeading = false;
var employees = sheet.ReadRows<Employee>();
```

### Column selectors

To choose columns in code, implement `IColumnSelector`. It's called once for each sheet (or region), with the sheet and the heading of the rows being read, and returns the positions of the columns to read, as `ExcelHeading.GetColumnIndex` gives them:

```csharp
public class PrefixSelector(string prefix) : IColumnSelector
{
    public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context)
        => context.Heading!.ColumnNames
            .Where(name => name.StartsWith(prefix))
            .Select(context.Heading.GetColumnIndex)
            .ToArray();

    // Used in error messages: Could not read value for member "Bonuses" for columns starting with "Bonus_".
    public string Describe(ColumnSelectorContext context) => $"columns starting with \"{prefix}\"";
}

public class Employee
{
    // A single value reads the first column chosen.
    [ExcelColumnMatching(typeof(PrefixSelector), ConstructorArguments = ["Salary"])]
    public decimal Salary { get; set; }

    // A collection or dictionary reads every column chosen, in order.
    [ExcelColumnsMatching(typeof(PrefixSelector), ConstructorArguments = ["Bonus_"])]
    public decimal[] Bonuses { get; set; }
}

// Or with a class map
importer.Configuration.RegisterClassMap<Employee>(c => c.Map(e => e.Bonuses).WithColumns(new PrefixSelector("Bonus_")));
```

If no column is chosen, the member is missing, as when a named column isn't found. When reading a region, chosen columns outside it are ignored.

### Missing and ignored columns

A missing column throws an `ExcelMappingException` that names it and suggests a similarly named column. Make a member optional to leave it unset instead (it keeps its initial value), or ignore it to never read it:

```csharp
public class Employee
{
    public string Name { get; set; }

    [ExcelOptional]                // or .MakeOptional()
    public string MiddleName { get; set; }

    [ExcelIgnore]                  // in a class map, just don't map it
    public DateTime ImportedAt { get; set; }
}
```

To check every column before reading any rows, see [Checking columns before reading](errors.md#checking-columns-before-reading).

### Row numbers

Set a member to the row's number, for example to report problems back to whoever filled in the spreadsheet:

```csharp
public class Employee
{
    [ExcelRowNumber]               // or c.MapRowNumber(e => e.Row)
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

Row numbers are one-based and count every row in the sheet, as Excel shows them: the heading, rows above it, and blank rows skipped by `SkipBlankLines`. So the first row under a heading in row 1 is row 2, matching the rows in error messages. The member doesn't read a column, and can be any numeric type. For a positional record, use `[property: ExcelRowNumber]`.

### Extra columns

Keep the columns nothing else reads, keyed by column name:

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

// Or with a class map, which also lets you configure the values
importer.Configuration.RegisterClassMap<Employee>(c =>
{
    c.Map(e => e.Name);
    c.Map(e => e.Age);
    c.MapExtraColumns(e => e.Other).WithValueMap(v => v.WithTrim());
});
```

A column is extra if no other member of the type being read reads it, including members of nested objects at any depth and members with column attributes or separators. Members with `[ExcelIgnore]`, and members a class map doesn't map, don't read a column, so their columns are extra. Columns with an empty name are skipped, and the sheet must have a heading. The member can be any dictionary with string keys, such as `Dictionary<string, string>`, `Dictionary<string, int>`, `IReadOnlyDictionary<string, object>` (values as Excel stores them) or `ExpandoObject`.

## Converting values yourself

The built-in conversion for each type is described in [Types](types.md). To change it:

| To | Attribute | Fluent method |
|---|---|---|
| Map fixed text to values | `[ExcelMappingDictionary("FT", Status.FullTime)]` | `.WithMapping(dictionary)` |
| Convert with your own code | `[ExcelConverter(typeof(Converter))]` | `.WithConverter(value => ...)` or `.WithConverter(converter)` |
| Change cells before they're converted | `[ExcelTrimString]`, `[ExcelTransformer(typeof(Transformer))]` | `.WithTrim()`, `.WithTransformers(...)` |
| Read the text as Excel displays it | `[ExcelPreserveFormatting]` | `.MakePreserveFormatting()` |

### Value mapping

Map fixed text, such as codes and abbreviations, to values:

| Name          | Size | Status |
|---------------|------|--------|
| Alice Johnson | L    | FT     |
| Bob Smith     | M    | C      |

```csharp
public enum TShirtSize { Small, Medium, Large }
public enum Status { FullTime, PartTime, Contract }

public class Employee
{
    public string Name { get; set; }

    [ExcelMappingDictionary("S", TShirtSize.Small)]
    [ExcelMappingDictionary("M", TShirtSize.Medium)]
    [ExcelMappingDictionary("L", TShirtSize.Large)]
    public TShirtSize Size { get; set; }

    [ExcelMappingDictionary("FT", Status.FullTime)]
    [ExcelMappingDictionary("PT", Status.PartTime)]
    [ExcelMappingDictionary("C", Status.Contract)]
    [ExcelMappingDictionaryComparer(StringComparison.OrdinalIgnoreCase)]
    public Status Status { get; set; }
}

// Or with a class map
importer.Configuration.RegisterClassMap<Employee>(c => c.Map(e => e.Status).WithMapping(new Dictionary<string, Status>
{
    ["FT"] = Status.FullTime,
    ["PT"] = Status.PartTime,
    ["C"] = Status.Contract,
}, StringComparer.OrdinalIgnoreCase));
```

Text is matched case-sensitively unless you give a comparer. Text that isn't in the dictionary is converted as usual (so `"Contract"` still maps to `Status.Contract`); to make it invalid instead, use `[ExcelMappingDictionaryBehavior(MappingDictionaryBehavior.Required)]` or `.WithMapping(dictionary, behavior: MappingDictionaryBehavior.Required)`.

### Converters

`WithConverter(value => ...)` converts a cell's text with a delegate:

```csharp
importer.Configuration.RegisterClassMap<Employee>(c => c.Map(e => e.Score).WithConverter(value => float.Parse(value!.TrimEnd('%')) / 100f));
```

For more control, implement `ICellConverter<T>`. A converter gets the cell's value as Excel stores it and where the cell is, and is set with `[ExcelConverter]` or `.WithConverter(converter)`:

```csharp
public class ThousandsConverter : ICellConverter<int>
{
    public CellConversionResult<int> Convert(CellValue value, CellContext context)
    {
        // Number cells from Excel files arrive as numbers, not text.
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
    public int Quantity { get; set; }    // "2k" → 2000

    [ExcelConverter(typeof(ThousandsConverter))]
    public int? Reorder { get; set; }    // a converter for T also converts T?

    [ExcelSeparators(';')]
    [ExcelConverter(typeof(ThousandsConverter))]
    public int[] Batches { get; set; }   // converts each element
}
```

- **A converter replaces the built-in conversion** for the member's type, so it's used for every cell that isn't empty. Empty cells use the member's empty fallback.
- **`CellValue`**: `Kind` (`Empty`, `Text`, `Number`, `Boolean`, `DateTime`, `TimeSpan` or `Other`), `Text`, `Number`, `Boolean`, `DateTime`, `TimeSpan` and `Value`. CSV cells are always text.
- **`CellContext`**: `Sheet`, the `Heading` of the rows being read (a region's, when reading one), `RowIndex`, `ColumnIndex`, `CellAddress` (such as `"B11"`), the `Member` being mapped, and the `FormatProvider` to parse with (the importer's [culture](types.md#culture), if set). To report a cell as an error, throw `new ExcelMappingException(message, context)`: the message gives the cell's address and column.
- **`CellConversionResult<T>`**: `Success(value)`; `Invalid()`, `Invalid(exception)` or `Invalid(() => exception)`, which uses the invalid fallback; or `NotHandled`, which lets the next converter try. Several `[ExcelConverter]` attributes are tried in order, and if none converts the value, it's invalid.
- **Converters are shared** between rows and threads, so they mustn't change their own state. `[ExcelConverter]` takes `ConstructorArguments` for converters with parameters.

### Transformers

A transformer changes a cell's value before it's converted. `[ExcelTrimString]` and `.WithTrim()` trim text; implement `ICellTransformer` for anything else:

```csharp
public class UpperCaseTransformer : ICellTransformer
{
    public CellValue Transform(CellValue value, CellContext context)
        => value.Kind == CellKind.Text ? new CellValue(value.Text!.ToUpperInvariant()) : value;
}

// A transformer can return a typed value: "1 234,5" becomes the number 1234.5.
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
    [ExcelTransformer(typeof(UpperCaseTransformer))]   // or .WithTransformers(new UpperCaseTransformer())
    public string Name { get; set; }

    [ExcelTrimString]
    [ExcelTransformer(typeof(FrenchNumberTransformer))]
    public double Salary { get; set; }
}
```

Transformers run in order, before the empty check, so a transformer that returns an empty value (`new CellValue((string?)null)`) makes the cell use the empty fallback. Trimming only changes text: cells Excel stores as numbers or dates keep their type.

### Preserving formatting

By default, a cell's value is read as Excel stores it, so a number formatted as `00123` or `$95,000` is the number 123 or 95000. Preserve formatting to read the text Excel displays instead:

| Employee ID | Salary  |
|-------------|---------|
| 00123       | $95,000 |

```csharp
public class Employee
{
    [ExcelPreserveFormatting]
    public string EmployeeId { get; set; }  // "00123"

    [ExcelPreserveFormatting]
    public string Salary { get; set; }      // "$95,000"
}
```

## Empty and invalid cells

By default:

- **An empty cell** gives `null` for reference types and nullable types, and an error for other value types (such as `int` and `DateTime`).
- **An invalid cell**, one that can't be converted to the member's type (such as `"N/A"` for an `int`), is an error.

An error throws an `ExcelMappingException` that names the cell, or is collected by [`ReadRowResults`](errors.md#collecting-errors). Give a fallback to use a value instead:

| Name          | Age |
|---------------|-----|
| Alice Johnson |     |
| Bob Smith     | N/A |

```csharp
public class Employee
{
    public string Name { get; set; }

    [ExcelDefaultValue(0)]     // empty cells; or .WithEmptyFallback(0)
    [ExcelInvalidValue(-1)]    // invalid cells; or .WithInvalidFallback(-1)
    public int Age { get; set; }
}
```

`.WithFallback(value)` sets both. `.WithThrowingEmptyFallback()`, `.WithThrowingInvalidFallback()` and `.WithThrowingFallback()` make empty or invalid cells errors, for example for a nullable member that mustn't be empty. To give every empty cell of a class map its type's default value, pass `EmptyValueStrategy.DefaultValue` to the `ExcelClassMap<T>` constructor.

### Computed fallbacks

To compute the value, implement `ICellFallback<T>` and set it with `[ExcelEmptyFallback]`/`[ExcelInvalidFallback]` or `.WithEmptyFallback(fallback)`/`.WithInvalidFallback(fallback)`:

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
        throw new ExcelMappingException($"Invalid age \"{value.Text}\".", context, exception);
    }
}

public class Employee
{
    [ExcelInvalidFallback(typeof(AgeFromBirthYearFallback))]
    public int Age { get; set; }
}
```

The fallback gets the cell's value (after any transformers), where it is, and why it's invalid (`null` for empty cells). A fallback for `T` also works for `T?` members, and both attributes take `ConstructorArguments`.

## Validation

Turn on `ValidateDataAnnotations` to check each object with its data annotations, such as `[Required]`, `[Range]` and `[StringLength]`, after it's mapped:

```csharp
using System.ComponentModel.DataAnnotations;

public class Employee
{
    public string Name { get; set; }

    [Range(1, 100)]
    public int Percentage { get; set; }
}

importer.Configuration.ValidateDataAnnotations = true;

try
{
    var employees = sheet.ReadRows<Employee>().ToArray();
}
catch (ValidationException ex)
{
    Console.WriteLine(ex.Message); // "The field Percentage must be between 1 and 100."
}
```

`ReadRows` throws the first `ValidationException`. `ReadRowResults` reports every failure of every row as an error instead.
