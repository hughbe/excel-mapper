# Reading Files

Opening workbooks and CSV files, choosing sheets, reading rows, and telling ExcelMapper where the heading and data are. See the [documentation index](../README.md#documentation) for other topics.

## Opening a file

`ExcelImporter.ReadRows<T>(path)` reads every row of a file's first sheet in one call. For anything else, create an importer:

```csharp
new ExcelImporter("data.xlsx");                          // a path
new ExcelImporter(stream);                               // a stream
new ExcelImporter("data.csv", ExcelImporterFileType.Csv); // a CSV file
new ExcelImporter(reader);                               // an IExcelDataReader you created
```

`.xlsx`, `.xls` and `.xlsb` workbooks are detected automatically. CSV files must be opened with `ExcelImporterFileType.Csv` (otherwise ExcelDataReader throws a `HeaderException`); their separator (comma, semicolon, tab, pipe or `#`) is detected automatically. For passwords, a specific CSV separator or a fallback encoding, see [Reader options](../README.md#reader-options).

To configure ExcelDataReader yourself, create the `IExcelDataReader` and pass it to `new ExcelImporter(reader)`. Don't move the reader afterwards: the importer's sheets share it.

## Choosing a sheet

```csharp
var first = importer.ReadSheet();          // the next sheet
var sales = importer.ReadSheet("Sales");   // by name, ignoring case
var third = importer.ReadSheet(2);         // by zero-based index

foreach (var sheet in importer.ReadSheets())
{
    Console.WriteLine($"{sheet.Name}: {sheet.Visibility}"); // Visible, Hidden or VeryHidden
}
```

Each `ReadSheet` throws an `ExcelMappingException` if there's no such sheet; `TryReadSheet(out sheet)`, `TryReadSheet(name, out sheet)` and `TryReadSheet(index, out sheet)` return `false` instead. `importer.NumberOfSheets` is the number of sheets.

You can read several sheets at once, including several `ExcelSheet` objects for the same worksheet: each keeps its own position. The file is read forward only, so going back to an earlier row re-reads it from the start. Reading one sheet's rows at a time is fastest; alternating between sheets row by row works, but re-reads the file each time.

## Reading rows

```csharp
foreach (var employee in sheet.ReadRows<Employee>())   // mapped lazily as you enumerate
{
    Console.WriteLine(employee.Name);
}

var first = sheet.ReadRow<Employee>();                // the next row; throws if there are none
if (sheet.TryReadRow<Employee>(out var next)) { }     // or returns false

var block = sheet.ReadRows<Employee>(startIndex: 5, count: 10);
```

- Every read method reads the heading first if the sheet has one and it hasn't been read, so the heading row is never mapped as data.
- `ReadRows(startIndex, count)` reads rows by their zero-based index in the sheet. With the heading in the first row (index 0), the first data row is index 1. It throws if the rows don't exist.
- To report every invalid cell instead of stopping at the first, use `ReadRowResults<T>()` (see [Collecting errors](errors.md#collecting-errors)).

A blank row is mapped like any other row: reference and nullable members get `null`, and other members use their empty fallback, or throw. Set `importer.Configuration.SkipBlankLines = true` to skip blank rows instead. Empty lines at the end of a CSV file are always ignored.

## Where the heading is

By default the first row is the heading, and every row after it is data.

**No heading.** Set `HasHeading` to `false`, and read columns by position:

|               |             |
|---------------|-------------|
| Alice Johnson | Engineering |
| Bob Smith     | Marketing   |

```csharp
public record Employee(
    [property: ExcelColumnIndex(0)] string Name,
    [property: ExcelColumnIndex(1)] string Department);

var sheet = importer.ReadSheet();
sheet.HasHeading = false;
var employees = sheet.ReadRows<Employee>();
var ids = sheet.ReadRows<int>();             // a single value reads the first column
```

Without a heading, a type whose members are matched by column name can't be mapped: map its members by position (with attributes or `.WithColumnIndex(...)`).

**A heading in another row.** Set `HeadingIndex` to its zero-based row:

|                      |             |
|----------------------|-------------|
| Employee Report 2025 |             |
|                      |             |
| Name                 | Department  |
| Alice Johnson        | Engineering |

```csharp
var sheet = importer.ReadSheet();
sheet.HeadingIndex = 2;   // the third row
var employees = sheet.ReadRows<Employee>();
```

### Finding the heading

If the rows above the heading vary between files, let ExcelMapper find it:

```csharp
var sheet = importer.ReadSheet();
sheet.FindHeading<Employee>();   // finds the "Name, Department" row
var employees = sheet.ReadRows<Employee>();
```

`FindHeading<T>()` uses the first row, among the first 50, that has every column `T` needs by name. Blank rows are skipped, and so are rows with only some of the columns, such as a title that happens to say "Name". It sets `HeadingIndex`, and reading continues from the row after the heading.

- Columns found by position, optional columns and dictionaries of all columns don't identify the heading on their own, but optional columns count towards a match.
- `FindHeading<T>(maxRows)` searches more or fewer rows. The search starts at `HeadingIndex`, so set that first to skip rows you know aren't the heading.
- If no row matches, `FindHeading` throws an `ExcelMappingException` naming the closest row and the columns it's missing. `TryFindHeading<T>(out var heading)` returns `false` instead and leaves the sheet as it was.

## Reading part of a sheet

A region is a block of a sheet with its own heading, for reading part of a sheet or several tables on one sheet:

|   | A     | B      | C | D     | E    |
|---|-------|--------|---|-------|------|
| 1 | Sales |        |   | Costs |      |
| 2 | Name  | Amount |   | Name  | Cost |
| 3 | a     | 1      |   | x     | 10   |
| 4 | b     | 2      |   | y     | 20   |

```csharp
var sheet = importer.ReadSheet();
var sales = sheet.Region("A2:B4").ReadRows<Sale>();
var costs = sheet.Region("D2:E4").ReadRows<Cost>();
```

A region's first row is its heading. Change that with:

- `WithHeadingAt(rowIndex)`: the heading is another row of the sheet, inside or above the region, and the data is the region's rows after it. `sheet.Region(new ExcelRange(100..200, Range.All)).WithHeadingAt(0)` reads rows 101 to 200 with the sheet's first row as their heading.
- `WithoutHeading()`: every row is data, and columns are found by position.
- `WithHeadingSearch()` or `WithHeadingSearch(maxRows)`: the heading is found as by [`FindHeading`](#finding-the-heading), within the region.

Regions have `ReadRows<T>()`, `ReadRowResults<T>()`, `ValidateHeading<T>()` and `TryValidateHeading<T>()`.

- **Regions are immutable**, and each read starts at the top of the region. Reading one doesn't use or change the sheet's heading, `DataRange` or position, so a region can be read any number of times, alongside other regions and the sheet.
- **Positions count from the region's first column.** In region `B2:C3`, `[ExcelColumnIndex(0)]` and `[ExcelColumnLetter("A")]` are column B, so one type can map tables of the same shape anywhere on a sheet. Error messages give cell addresses in the sheet.
- **Blank rows** are mapped like other rows unless `SkipBlankLines` is set. A region of columns `A:B` includes rows whose values are all in other columns.

## Restricting the sheet's data

`DataRange` limits the rows and columns the sheet itself reads. To read several blocks, use [regions](#reading-part-of-a-sheet) instead.

|                      |             |          |
|----------------------|-------------|----------|
| Employee Report 2025 |             |          |
| Name                 | Department  | Obsolete |
| Alice Johnson        | Engineering | Old      |
| Bob Smith            | Marketing   | Old      |
|                      |             |          |
| Generated 2025-01-01 |             |          |

```csharp
var sheet = importer.ReadSheet();
sheet.DataRange = ExcelRange.Parse("A2:B4");   // the heading and two rows, columns A and B
var employees = sheet.ReadRows<Employee>();
```

With a heading, the heading is the first row of the range. Set `DataRange` before reading the heading or any rows. Positions in `[ExcelColumnIndex]` and similar stay the sheet's, even when the range doesn't start at column A.

Ranges can be written in several ways:

```csharp
ExcelRange.Parse("B2:D10");                    // an Excel address, also "$B$2:$D$10"
ExcelRange.Parse("A:C");                       // whole columns
ExcelRange.Parse("5:20");                      // whole rows
new ExcelRange(rowStart: 1, rowEnd: 99, columnStart: 0, columnEnd: null);   // zero-based, inclusive; null for no limit
new ExcelRange(..^1, Range.All);               // System.Range: every row but the last
```

Addresses must be within Excel's limits (rows 1 to 1,048,576, columns A to XFD): for larger CSV files, use the index or `Range` constructors. Row ranges counted from the end (`^1`) read a CSV file one extra time to count its rows. `default(ExcelRange)` is every row and column.

## Limits

To protect against files with an excessive number of columns, a sheet with more than `importer.Configuration.MaxColumnsPerSheet` columns (10,000 by default) throws an `ExcelMappingException` before anything is allocated for them. Excel itself allows 16,384. Raise the limit only for files you trust.
