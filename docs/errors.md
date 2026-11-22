# Errors and Troubleshooting

What errors look like, how to check a file before reading it, how to report every problem in a file at once, and fixes for common problems. See the [documentation index](../README.md#documentation) for other topics.

## Error messages

When a sheet, column or cell can't be mapped, ExcelMapper throws an `ExcelMappingException` that says where the problem is, as Excel shows it: a cell address when the column is known, otherwise a one-based row number.

```csharp
try
{
    var employees = sheet.ReadRows<Employee>().ToArray();
}
catch (ExcelMappingException ex)
{
    // Cannot assign "abc" to member "YearsOfService" of type "System.Int32" in column "YearsOfService" at cell C3 in sheet "Employees".
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
- `Errors` lists every problem in the row: invalid or empty values, missing columns, exceptions thrown by a constructor, and failed data annotations when `ValidateDataAnnotations` is on.
- `RowIndex` is the row's zero-based index in the sheet.
- Problems with the mapping itself, such as a type that can't be auto-mapped, still throw.
- A missing column is reported on every row, so call `ValidateHeading` first to report it once.

`ReadRowResults` costs the same as `ReadRows` for valid rows.

## Troubleshooting

**`Could not read value for member "FirstName" for column "FirstName"`.** The sheet has no column with that name. Check the message for a suggested column, then:
- rename the member's column with `[ExcelColumnName]`, or accept several names (see [Choosing columns](mapping.md#choosing-columns));
- set `importer.Configuration.IgnoreSeparatorsInColumnNames = true` to match "First Name", "first_name" and so on;
- make the member optional if the column isn't always there ([Missing and ignored columns](mapping.md#missing-and-ignored-columns));
- check where the heading is: see [Where the heading is](reading.md#where-the-heading-is) and [Finding the heading](reading.md#finding-the-heading).

**`Cannot assign "" to member ...` for an empty cell.** The member's type can't be empty, such as `int` or `DateTime`. Make it nullable (`int?`), or give it an empty fallback ([Empty and invalid cells](mapping.md#empty-and-invalid-cells)).

**`Cannot assign "N/A" to member ...` for text that isn't a valid value.** Give the member an invalid fallback, map the text with [value mapping](mapping.md#value-mapping), or [convert it yourself](mapping.md#converters).

**A fractional number such as `2.7` is invalid for an `int` member.** Excel stores every number as a `double`, and ExcelMapper doesn't silently drop or round the fraction (3.x did). Map to `decimal` or `double`, give the member an invalid fallback, or round it in a converter: `.WithConverter(value => (int)Math.Round(double.Parse(value)))`.

**Dates from a CSV file are invalid, or the day and month are swapped.** Text is parsed with the importer's culture, invariant by default (month first). Set `importer.Configuration.Culture`, or give the member its formats ([Dates and times](types.md#dates-and-times)).

**`DateTimeOffset` values have a different offset on another machine.** Excel dates have no time zone, so they get the local offset: set one with `[ExcelOffset]` ([Offsets](types.md#offsets)).

**The heading row, or a title, is mapped as data.** The heading isn't in the first row: set `HeadingIndex` or call `FindHeading<T>()` ([Where the heading is](reading.md#where-the-heading-is)).

**Blank rows cause errors.** Set `importer.Configuration.SkipBlankLines = true`.

**`Circular reference detected: type "Person" references itself through its members`.** A type that contains itself can't be auto-mapped. Mark the member `[ExcelIgnore]`, or map the type with a class map.

**`No data is available for encoding 1252`** when you create an `IExcelDataReader` yourself. `ExcelImporter` registers code page encodings, but only once one has been created. Register them first: `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)`.

**`HeaderException` when opening a CSV file.** Open it with `ExcelImporterFileType.Csv` ([Opening a file](reading.md#opening-a-file)).

**`Cannot auto-map type` in a trimmed app.** The type's members were removed by the trimmer: see [Trimming](trimming.md).
