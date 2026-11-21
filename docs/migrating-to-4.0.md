# Migrating to 4.0

ExcelMapper 4.0 fixes several long-standing bugs and makes mapping more predictable. Most code keeps working unchanged: the public API is compatible with 3.0.3 apart from one type the netstandard builds exposed by mistake (checked by package validation), but some **behaviour** changed. This page lists what to check, most commonly affected first. See the [changelog](../CHANGELOG.md) for everything that's new.

## Checklist

1. [`object` members now receive typed values](#object-members-receive-typed-values)
2. [Fractional numbers can't be mapped to integers](#fractional-numbers-cant-be-mapped-to-integers)
3. [`ReadRow` reads the heading first](#readrow-reads-the-heading-first)
4. [Error messages use cell addresses](#error-messages-use-cell-addresses)
5. [Sheet names are case-insensitive](#sheet-names-are-case-insensitive)
6. [Smaller changes](#smaller-changes)
7. [You can remove the code page registration](#you-can-remove-the-code-page-registration)

## `object` members receive typed values

**Who's affected:** code that maps to `object` or `IConvertible` members, `Dictionary<string, object>`, `ExpandoObject`/`dynamic`, or untyped collections such as `ArrayList`, and then treats the values as strings.

**Before:** every value was the cell's text: a number `2.5` arrived as `"2.5"` (or `"2,5"`, depending on the machine's culture), a date as culture-formatted text, a boolean as `"True"`.

**Now:** values arrive as Excel stores them: `double` for numbers, `DateTime` for dates, `bool` for booleans, `TimeSpan` for durations and `string` for text.

**What to do:**

```csharp
// Before: worked because every value was a string
var amount = (string)row["Amount"];

// After: use the typed value...
var amount = (double)row["Amount"];

// ...or convert explicitly, choosing the culture
var text = Convert.ToString(row["Amount"], CultureInfo.InvariantCulture);
```

To keep receiving text for a member, map it to `string`, or preserve formatting (`[ExcelPreserveFormatting]` or `.MakePreserveFormatting()`) to get the text as displayed in Excel. CSV files are unaffected: every CSV cell is text.

## Fractional numbers can't be mapped to integers

**Who's affected:** code that maps numeric cells holding fractions (such as `2.7`) to `int`, `long`, `byte` and other integer types.

**Before:** the fraction was silently dropped: truncated to `2` on .NET 8 and rounded to `3` on .NET Framework and other netstandard runtimes.

**Now:** the cell is invalid, the same as the text `"2.7"`. Whole numbers such as `2.0` still map.

**What to do:** if fractions are expected, map to `decimal` or `double`, or decide what an invalid value becomes:

```csharp
[ExcelInvalidValue(0)]
public int Quantity { get; set; }

// Or round explicitly
Map(e => e.Quantity)
    .WithConverter(value => (int)Math.Round(double.Parse(value, CultureInfo.InvariantCulture)));
```

To find affected rows without stopping at the first one, use `sheet.ReadRowResults<T>()`.

## `ReadRow` reads the heading first

**Who's affected:** code that calls `ReadRow`/`TryReadRow` before `ReadHeading()` on a sheet that has a heading.

**Before:** the first `ReadRow` call mapped the heading row as if it were data.

**Now:** `ReadRow` and `TryReadRow` read the heading first, as `ReadRows` always has, so the first call returns the first data row.

**What to do:** remove any code that skipped the heading row by hand. If the sheet really has no heading, say so:

```csharp
var sheet = importer.ReadSheet();
sheet.HasHeading = false;
```

## Error messages use cell addresses

**Who's affected:** code or tests that compare `ExcelMappingException.Message` text.

**Before:** `... on row 10 in sheet "Sheet1".` (a zero-based row index).

**Now:** `... at cell B11 in sheet "Sheet1".` when the column is known, otherwise `... on row 11 ...` (one-based, as Excel shows it).

**What to do:** use the exception's properties instead of parsing the message:

```csharp
catch (ExcelMappingException ex)
{
    Console.WriteLine(ex.CellAddress); // "B11"
    Console.WriteLine(ex.ColumnName);  // "Salary"
    Console.WriteLine(ex.Member?.Name);
    Console.WriteLine(ex.RowIndex);    // 10 (zero-based, unchanged)
}
```

## Sheet names are case-insensitive

**Who's affected:** workbooks with sheets whose names differ only by case, or code relying on `TryReadSheet("summary", ...)` failing for a sheet named "Summary".

**Now:** `ReadSheet(name)` and `TryReadSheet(name, ...)` ignore case, as Excel does (Excel doesn't allow two sheets whose names differ only by case).

## Smaller changes

- **`ExcelRange`**: `default(ExcelRange)`, and the value `TryParse` returns on failure, now mean every row and column (previously an empty range). `Parse` rejects addresses Excel doesn't accept: non-ASCII letters, signs (`1:+5`), columns past `XFD` and rows past 1,048,576. Use the `Range` or index constructors for larger CSV files.
- **Column offsets**: with a `DataRange` that doesn't start at column A, `ExcelHeading.GetColumnIndex` and `GetColumnName` use the column's position in the sheet (as cell readers always did). Before 4.0, values were read from the wrong columns in this case.
- **Split values**: `"a,b,"` split into an array or list now includes the trailing empty value, which maps through the empty-value fallback (`null` for strings; an error for non-nullable numbers unless you add a fallback). An empty cell gives an empty collection.
- **Registering after auto-mapping**: `RegisterClassMap<T>` now replaces a map that auto-mapping created for `T` (for example by an earlier `ReadRows<T>()`) instead of throwing. Registering twice yourself still throws.
- **Constructor exceptions**: an exception thrown by a constructor called during mapping (for records and constructor-only types) is wrapped in an `ExcelMappingException` with the row, with the original as `InnerException`.
- **Trimming**: `[ExcelTrimString]`/`.WithTrim()` no longer turn date and number cells into text, so they keep their type.
- **`bool`**: "yes" and "no" (any case) are now valid, where they were invalid.
- **Dates written as text**: without set formats, ISO 8601 (`2024-03-04`) and the culture's short date, short date and time, and time patterns are now accepted, where only the culture's full date and time (`"G"`, or `"d"`/`"t"` for `DateOnly`/`TimeOnly`) was before. Values that were invalid, and so used an `[ExcelInvalidValue]` or other fallback, may now map. To accept only specific formats, set them with `[ExcelFormats]` or `.WithFormats(...)`.
- **Dates stored as numbers**: a number in a cell without a date format now maps to date and time types (as an Excel serial date) instead of being invalid.
- **Number parsing is the same on every build**: on .NET Framework and other netstandard runtimes, `[ExcelNumberStyle]` now takes effect (it was silently ignored), and integers accept thousands separators such as `"1,234"`, as on .NET 8. On every build, floating-point types accept exponents such as `"1E5"`. These only accept more input than before, except where an `[ExcelNumberStyle]` attribute on netstandard now restricts what's accepted.
- **netstandard builds**: argument exceptions now include the parameter name, and `null` passed to string arguments throws `ArgumentNullException` (previously `ArgumentException`), matching the .NET 8 build. The public `CallerArgumentExpressionAttribute` type that 3.x accidentally exposed is gone.

## You can remove the code page registration

ExcelMapper now registers code page encodings itself, so this line is no longer needed (it's harmless to keep):

```csharp
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
```

On .NET Framework, .NET Standard and older .NET apps, the package now brings in `System.Text.Encoding.CodePages`.

## Features that need .NET 8

The package contains `net8.0`, `netstandard2.1` and `netstandard2.0` builds. If your app targets .NET Framework or .NET before 8, some features aren't available; this is unchanged from 3.x:

- mapping `DateOnly`, `TimeOnly`, `Half`, `Int128`, `UInt128`, `nint`, `nuint`, `BigInteger`, `Complex` and other `IParsable<T>` / `INumberBase<T>` types
- `.WithNumberStyle(...)` and `.WithFormatProvider(...)` for numeric members; on the netstandard builds use `[ExcelNumberStyle]` and `importer.Configuration.Culture` instead
- `.WithFormats(...)` for `DateOnly` and `TimeOnly`, and `IReadOnlySet<T>` members

## New in 4.0

Worth a look while upgrading:

- [Collect every error in a file](errors.md#collecting-errors) with `ReadRowResults`, and [check columns up front](errors.md#checking-columns-before-reading) with `ValidateHeading`.
- [Find the heading](reading.md#finding-the-heading) of sheets with titles or notes above it with `FindHeading`.
- [Records and constructor-only types](mapping.md#records-and-constructors) map without a parameterless constructor.
- Map by [column letter](mapping.md#column-index-mapping), choose a [culture](mapping.md#culture-for-the-whole-importer) for the whole importer, and set the [offset of `DateTimeOffset` values](types.md#date-and-time-types).
- `ExcelImporter.ReadRows<T>("file.xlsx")` reads a file in one line.
