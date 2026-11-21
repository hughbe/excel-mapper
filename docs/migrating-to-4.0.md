# Migrating to 4.0

ExcelMapper 4.0 fixes several long-standing bugs and makes mapping more predictable. Most code keeps working unchanged: reading rows, attributes, class maps and the fluent API are source compatible, but some **behaviour** changed, and the low-level extension types are replaced by simpler ones (see [Extension types](#extension-types)). This page lists what to check, most commonly affected first. See the [changelog](../CHANGELOG.md) for everything that's new.

## Checklist

1. [`object` members now receive typed values](#object-members-receive-typed-values)
2. [Fractional numbers can't be mapped to integers](#fractional-numbers-cant-be-mapped-to-integers)
3. [`WithConverter` replaces the built-in conversion](#withconverter-replaces-the-built-in-conversion)
4. [`ReadRow` reads the heading first](#readrow-reads-the-heading-first)
5. [Error messages use cell addresses](#error-messages-use-cell-addresses)
6. [Sheet names are case-insensitive](#sheet-names-are-case-insensitive)
7. [Extension types](#extension-types)
8. [Smaller changes](#smaller-changes)
9. [You can remove the code page registration](#you-can-remove-the-code-page-registration)

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

## `WithConverter` replaces the built-in conversion

**Who's affected:** code that calls `WithConverter` on a member whose type ExcelMapper converts itself (numbers, dates, enums, `bool`, `Guid` and so on), expecting the converter to only handle values the built-in conversion can't.

**Before:** the converter was added after the built-in conversion, so whether it ran depended on the member's type. For a number, the converter was skipped whenever the cell was a valid number. For a string, it was used for every cell:

```csharp
Map(e => e.Quantity).WithConverter(value => int.Parse(value) * 2);  // "3" gave 3, not 6
Map(e => e.Name).WithConverter(value => value.ToUpperInvariant());  // "a" gave "A"
```

**Now:** the converter replaces the built-in conversion, and any earlier converter, so it's used for every cell that isn't empty, whatever the type: `"3"` gives `6`. Empty cells still use the empty fallback, and a converter that throws uses the invalid fallback.

**What to do:** if the converter was meant as a fallback for values the built-in conversion can't handle, handle both in the converter:

```csharp
Map(e => e.Quantity).WithConverter(value =>
    int.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var quantity)
        ? quantity
        : ParseWords(value));
```

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

## Extension types

**Who's affected:** code that implements or creates ExcelMapper's low-level types directly, rather than using attributes, class maps and the fluent API.

4.0 has a smaller set of extension points, and the types they replace are internal:

| 3.x | 4.0 |
|---|---|
| Collection factories (`ExcelMapper.Factories`, `IEnumerableFactory<T>`, `IDictionaryFactory<TKey, TValue>`, `IMultidimensionalArrayFactory<T>`) and the collection and dictionary maps' constructors | ExcelMapper creates the collection for the member's type. Get a member's map from `ExcelClassMap<T>.Map(...)`. |
| Readers and reader factories (`ExcelMapper.Readers`, `ICellReader`, `ICellReaderFactory`, `ICellsReader`, `ICellsReaderFactory`), `IExcelColumnMatcher`, `WithReaderFactory(...)` and `WithColumnMatching(matcher)`/`WithColumnsMatching(matcher)` | [`IColumnSelector`](mapping.md#pattern-matching) with `WithColumns(selector)`, `[ExcelColumnMatching(typeof(...))]` or `[ExcelColumnsMatching(typeof(...))]`. |
| `IToOneMap`, `IOneToOneMap`, `IManyToOneMap` and the extension classes `IOneToOneMapExtensions`, `IManyToOneMapExtensions` and `IToOneMapExtensions`, `OneToOneMap<T>`'s constructor and the maps' `ReaderFactory` properties | The column methods (`WithColumnName`, `WithColumnNames`, `MakeOptional`, ...) are methods of the maps, so calls to them compile unchanged. Code that took an `IOneToOneMap` should take the map's own type. |
| `IValuePipelineExtensions`, `IValuePipeline<T>` as the receiver of the value methods, the collection maps' `IValuePipeline<T> Pipeline` and `WithElementMap`/`WithValueMap(Func<IValuePipeline<T>, IValuePipeline<T>>)`, `new ValuePipeline()` | `ValuePipelineExtensions`, whose methods extend `ValuePipeline<T>`. `OneToOneMap<T>` is a `ValuePipeline<T>`, and the collection maps' `Pipeline` is a `ValuePipeline<T>`, so calls and lambdas compile unchanged. Create a `ValuePipeline<T>` instead of a `ValuePipeline`. |
| `ICellMapper`, `CellMapperResult`, `ReadCellResult`, `WithMappers(...)`, `[ExcelMapper]`, `WithConverter(ConvertUsingMapperDelegate)` and the mappers in `ExcelMapper.Mappers` | An [`ICellConverter<T>`](mapping.md#custom-converters) with `WithConverter(converter)` or `[ExcelConverter(typeof(...))]`, or `WithConverter(value => ...)`. A converter returns `NotHandled` to let the built-in conversion try. |
| `Abstractions.ICellTransformer` and `TrimStringCellTransformer` | [`ICellTransformer`](mapping.md#custom-cell-transformers) (in the `ExcelMapper` namespace), and `WithTrim()` or `[ExcelTrimString]`. |
| `IFallbackItem`, `FixedValueFallback`, `ThrowFallback`, `WithFallbackItem`, `WithEmptyFallbackItem` and `WithInvalidFallbackItem` | [`ICellFallback<T>`](mapping.md#advanced-fallback-strategies) with `WithEmptyFallback(fallback)`, `WithInvalidFallback(fallback)`, `[ExcelEmptyFallback]` or `[ExcelInvalidFallback]`; `WithValueFallback`, `WithEmptyFallback(value)`, `WithThrowingFallback()` and the other fluent fallback methods. |
| A pipeline's `Mappers`, `Transformers`, `EmptyFallback` and `InvalidFallback` properties, and `OneToOneMap<T>.Pipeline` | The fluent methods above, which replace these. |
| `ExcelMapper.Mappers.MappingDictionaryMapperBehavior` | `ExcelMapper.MappingDictionaryMapperBehavior`. |
| `IMap`, `ExcelPropertyMap`, `ExcelClassMap.Properties`, the indexer maps, `AutoMapper.TryCreateClassMap` and `RegisterClassMap(Type, IMap)` | Map members with `ExcelClassMap<T>.Map(...)` (including `Map(p => p.Values[0])` for elements) and the fluent methods. To start from the auto-mapped class map, use `RegisterClassMap<T>(c => ...)`. `TryGetClassMap` returns an `ExcelClassMap`, and `RegisterClassMap(Type, ExcelClassMap)` registers a class map for another type. |

**What to do:** if you created a map with its constructor, or implemented `IMap` or added to `ExcelClassMap.Properties`, map the member with `Map(...)` in a class map instead. To control how values are created, use an [`ICellConverter<T>`](mapping.md#custom-converters) instead of an `ICellMapper`, and an `ICellTransformer` or `ICellFallback<T>` instead of the old transformer and fallback types. To choose columns in code, implement an [`IColumnSelector`](mapping.md#pattern-matching) instead of a reader factory or `IExcelColumnMatcher`.

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
- [Read part of a sheet](reading.md#reading-part-of-a-sheet), or several tables on one sheet, with `sheet.Region("B3:F20")`.
- [Records and constructor-only types](mapping.md#records-and-constructors) map without a parameterless constructor.
- Map by [column letter](mapping.md#column-index-mapping), choose a [culture](mapping.md#culture-for-the-whole-importer) for the whole importer, and set the [offset of `DateTimeOffset` values](types.md#date-and-time-types).
- `ExcelImporter.ReadRows<T>("file.xlsx")` reads a file in one line.
