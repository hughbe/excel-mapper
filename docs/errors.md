# Errors and Troubleshooting

What errors look like, how to check a file before reading it, how to report every problem in a file at once, and fixes for common problems. See the [documentation index](../README.md#documentation) for other topics.

## Error messages

When a sheet, column or cell can't be mapped, ExcelMapper throws an `ExcelMappingException` that says where the problem is, as Excel shows it: a cell address when the column is known, otherwise a one-based row number. Types are named as C# names them (`int`, `List<string>`), and a missing column's message lists the columns the heading has. A CSV file's sheet has no name, so its messages don't give one.

```csharp
try
{
    var employees = sheet.ReadRows<Employee>().ToArray();
}
catch (ExcelMappingException ex)
{
    // Cannot assign "abc" to member "YearsOfService" of type int in column "YearsOfService" at cell C3 in sheet "Employees".
    Console.WriteLine(ex.Message);

    Console.WriteLine(ex.CellAddress);      // "C3" (null if unknown)
    Console.WriteLine(ex.ColumnName);       // "YearsOfService" (null if unknown)
    Console.WriteLine(ex.Member?.Name);     // "YearsOfService"
    Console.WriteLine(ex.RowIndex);         // 2 (zero-based)
    Console.WriteLine(ex.InnerException);   // why the value was invalid, such as a FormatException
}
```

## Checking columns before reading

A missing column is otherwise reported when the first row is read, one column at a time. To check every column a type needs before reading any rows, validate the heading:

```csharp
var sheet = importer.ReadSheet();

// Throws an ExcelMappingException listing every missing column:
// Sheet "Employees" is missing columns needed to map type Employee: column "Salary" (member Salary); column "Zip" (member Address.Zip).
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

Optional members aren't required. Nested objects, collections and indexed members are checked too. If a column's name differs only by spaces, underscores, hyphens or dots, the problem suggests it: `column "Salary", did you mean "Sal-ary"?`. Validating reads the heading but no rows.

## Collecting errors

`ReadRows` throws at the first cell that can't be mapped. To check a whole file and report every problem, for example in an import screen, use `ReadRowResults`, which records every invalid cell in every row and keeps going:

```csharp
var results = sheet.ReadRowResults<Employee>().ToList();

var employees = results.Where(r => r.IsValid).Select(r => r.Value!).ToList();

foreach (var result in results.Where(r => !r.IsValid))
{
    foreach (var error in result.Errors)
    {
        Console.WriteLine($"{error.CellAddress}: {error.Message}");
    }
}
```

- `Value` is the mapped object, or `default` if the row has errors: partly mapped objects are never returned.
- `Errors` lists every problem in the row: invalid or empty values, exceptions thrown by a constructor, and failed data annotations when `ValidateDataAnnotations` is on.
- `RowIndex` is the row's zero-based index in the sheet.
- Problems with the sheet or the mapping itself throw when `ReadRowResults` is called, rather than being reported on every row: columns the type needs that the heading doesn't have (with the same message as `ValidateHeading`), and a type that can't be auto-mapped.

`ReadRowResults` costs the same as `ReadRows` for valid rows.

## Checking large files

`ReadRowResults` creates an `ExcelMappingException` for every invalid cell. To check a large file with many invalid cells without allocating for each, pass an `IExcelErrorSink` to `ReadRows`. It returns the rows without problems, and gives the sink each problem as it's found:

```csharp
public sealed class ErrorCounter : IExcelErrorSink
{
    public int Count { get; private set; }

    public void Add(in ExcelMappingError error)
    {
        Count++;
        if (Count <= 100)
        {
            // Creating the exception gives its message, and allocates.
            Console.WriteLine(error.ToException().Message);
        }
    }
}

var counter = new ErrorCounter();
foreach (var employee in sheet.ReadRows<Employee>(counter))
{
    Process(employee);
}

Console.WriteLine($"{counter.Count} problems");
```

- `ExcelMappingError` has the problem's `Kind` (`EmptyValue`, `InvalidValue`, `DataAnnotation` or `Exception`), `Sheet`, `RowIndex`, `ColumnIndex`, `ColumnName`, `Member` and the cell's `Value`. None of them allocate, except `Value.Text` for a cell that isn't text, such as a number.
- It's only valid during `Add`, as it refers to the cell while the reader is on its row: copy what you need, or call `ToException()`, which creates the exception `ReadRowResults` would collect.
- A row with a problem is skipped; a row can report several problems.
- As with `ReadRowResults`, columns the type needs that the heading doesn't have, and a type that can't be auto-mapped, throw when `ReadRows` is called.

## Troubleshooting

**`Can't find column "FirstName" for member "FirstName"`.** The heading has no column with that name: the message lists the columns it has, and suggests one whose name only differs by spaces or separators. Then:
- rename the member's column with `[ExcelColumnName]`, or accept several names (see [Choosing columns](mapping.md#choosing-columns));
- set `importer.Configuration.IgnoreSeparatorsInColumnNames = true` to match "First Name", "first_name" and so on;
- make the member optional if the column isn't always there ([Missing and ignored columns](mapping.md#missing-and-ignored-columns));
- check where the heading is: see [Where the heading is](reading.md#where-the-heading-is) and [Finding the heading](reading.md#finding-the-heading).

**`More than one column matches "Name" (columns A and B)`.** The heading has two columns with the member's name (ignoring case, and separators if `IgnoreSeparatorsInColumnNames` is set), so it isn't clear which to read. Choose one with `[ExcelColumnLetter]` or `.WithColumnLetter(...)`, or rename one of the columns.

**`The class map for type Order maps no members`.** Nothing in the type is mapped, so every row would be an empty object: its properties are read-only or have private setters, are all `[ExcelIgnore]`, or a class map registered for it has no `Map(...)` calls. Give the properties public setters or a constructor that takes them, or map them in a class map (`Map(p => p.Total)` works with a private setter).

**`The cell is empty, but member "Age" of type int can't be empty`.** The member's type can't be empty, such as `int` or `DateTime`. Make it nullable (`int?`), or give it an empty fallback ([Empty and invalid cells](mapping.md#empty-and-invalid-cells)).

**`Cannot assign "N/A" to member ...` for text that isn't a valid value.** For a number that's too large for its type, or has a fraction for an integer type, the message says so. Give the member an invalid fallback, map the text with [value mapping](mapping.md#value-mapping), or [convert it yourself](mapping.md#converters).

**A fractional number such as `2.7` is invalid for an `int` member.** Excel stores every number as a `double`, and ExcelMapper doesn't silently drop or round the fraction (3.x did). Map to `decimal` or `double`, give the member an invalid fallback, or round it in a converter: `.WithConverter(value => (int)Math.Round(double.Parse(value)))`.

**Dates from a CSV file are invalid, or the day and month are swapped.** Text is parsed with the importer's culture, invariant by default (month first). Set `importer.Configuration.Culture`, or give the member its formats ([Dates and times](types.md#dates-and-times)).

**`DateTimeOffset` values have a different offset on another machine.** Excel dates have no time zone, so they get the local offset: set one with `[ExcelOffset]` ([Offsets](types.md#offsets)).

**The heading row, or a title, is mapped as data.** The heading isn't in the first row: set `HeadingIndex` or call `FindHeading<T>()` ([Where the heading is](reading.md#where-the-heading-is)).

**Blank rows cause errors.** Set `importer.Configuration.SkipBlankLines = true`.

**`Circular reference detected: type "Person" references itself through its members`.** A type that contains itself can't be auto-mapped. Mark the member `[ExcelIgnore]`, or map the type with a class map.

**`No data is available for encoding 1252`** when you create an `IExcelDataReader` yourself. `ExcelImporter` registers code page encodings, but only once one has been created. Register them first: `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)`.

**`HeaderException` when opening a CSV file.** Open it with `ExcelImporterFileType.Csv` ([Opening a file](reading.md#opening-a-file)).

**`Cannot auto-map type` in a trimmed app.** The type's members were removed by the trimmer: see [Trimming](trimming.md).
