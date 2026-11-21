# Reading Files

Opening workbooks and CSV files, choosing sheets, reading rows, headings and data ranges. See the [documentation index](../README.md#documentation) for other topics.

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

You can read from several sheets at once, including several `ExcelSheet` objects for the same worksheet: each keeps its own position. The importer reads the file forward only, so going back to an earlier row of a worksheet re-reads it from the start. Reading one sheet's rows at a time (for example with `ToList()`) is fastest; alternating between two sheets row by row works but re-reads the file each time.

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

Without it, a blank row in the middle of a sheet is mapped like any other row: string and nullable members get `null`, and non-nullable members use their empty-value fallback (or throw). Empty lines at the end of a CSV file are always ignored.

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

#### Finding the heading

If the number of rows above the heading varies between files, let ExcelMapper find it:

```csharp
var sheet = importer.ReadSheet();
sheet.FindHeading<Employee>();  // Finds the "Name, Department" row

var employees = sheet.ReadRows<Employee>();
```

`FindHeading<T>()` reads the heading from the first row, among the first 50, that has every column `T` needs by name and at least one of its columns. Blank rows are skipped, and so are rows with only some of the required columns, such as a title that happens to say "Name". It sets `HeadingIndex`, and reading continues from the row after the heading.

- Columns found by index or letter, optional columns and dictionaries of all columns don't identify the heading row on their own; optional columns count towards a match.
- Search more or fewer rows with `FindHeading<T>(maxRows)`. The search starts at `HeadingIndex`, so set it first to skip rows you know aren't the heading.
- If no row matches, `FindHeading` throws an `ExcelMappingException` naming the closest row and the columns it's missing. `TryFindHeading<T>(out var heading)` returns `false` instead and leaves the sheet as it was.

### Reading Part of a Sheet

Read a block of a sheet, or several tables on one sheet, with regions:

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

- `WithHeadingAt(rowIndex)` - the heading is another row of the sheet, inside or above the region. The data is the region's rows after it, so `sheet.Region(new ExcelRange(100..200, Range.All)).WithHeadingAt(0)` reads rows 101 to 200 with the sheet's first row as their heading.
- `WithoutHeading()` - every row is data, and columns are found by index or letter.
- `WithHeadingSearch()` (or `WithHeadingSearch(maxRows)`) - the heading is found as by [`FindHeading`](#finding-the-heading), within the region.

Regions also have `ReadRowResults<T>()`, `ValidateHeading<T>()` and `TryValidateHeading<T>()`, which work as on a sheet.

- Regions are immutable, and each read starts at the top of the region. Reading a region doesn't use or change the sheet's own heading, `DataRange` or position, so a region can be read any number of times, alongside other regions and the sheet.
- A region's address is in the sheet, but columns chosen by position count from the region's first column: in region `B2:C3`, `[ExcelColumnIndex(0)]` and `[ExcelColumnLetter("A")]` are column B, and a position past the region's last column is a missing column. So one type can map tables of the same shape anywhere on a sheet. Error messages give cell addresses in the sheet. (On the sheet itself, positions are always the sheet's, even with `DataRange`.)
- Blank rows in a region are mapped like other rows unless `SkipBlankLines` is set. A region of columns `A:B` includes rows whose values are all in other columns.
- The file is read forward only. Reading one region at a time is as fast as reading the sheet; alternating between regions row by row re-reads the file each time.

### Restricting Data Range

Limit which rows and columns the sheet itself reads using `DataRange` (to read several blocks, use [regions](#reading-part-of-a-sheet)):

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
- Row ranges counted from the end (`^1`) of a CSV file read the file one extra time to count its rows, as the CSV reader's own row count leaves out empty lines. Empty lines at the end of the file aren't counted

**Common Use Cases:**

```csharp
// Skip header and footer rows
sheet.DataRange = new ExcelRange("A2:Z100");

// Process only specific columns
sheet.DataRange = new ExcelRange("B:E");  // Columns B, C, D, E

// Leave out a totals row at the bottom (rows counted from the end)
sheet.DataRange = new ExcelRange(..^1, Range.All);

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

## CSV Support

ExcelMapper supports CSV files in addition to Excel formats:

```csharp
// CSV files must be opened with ExcelImporterFileType.Csv
using var importer = new ExcelImporter("employees.csv", ExcelImporterFileType.Csv);
var sheet = importer.ReadSheet();
var employees = sheet.ReadRows<Employee>();
```

Comma, semicolon, tab, pipe and `#` separators are detected automatically. To use a specific separator or encoding, pass an `ExcelReaderConfiguration` (see [Reader options](../README.md#reader-options)).

**Supported file formats:**
- `.xlsx` - Excel 2007+ (Office Open XML)
- `.xls` - Excel 97-2003 (Binary Format)  
- `.xlsb` - Excel Binary Workbook
- `.csv` - Comma-separated values

**Note:** `new ExcelImporter(path)` only detects Excel formats. Opening a CSV file without `ExcelImporterFileType.Csv` throws a `HeaderException`.
