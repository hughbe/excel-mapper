# API Reference

The public types, attributes and fluent methods, grouped by what they do. Every type is in the `ExcelMapper` namespace. The package includes XML documentation, so each member's details also show in your editor. See the [documentation index](../README.md#documentation) for guides.

## Reading files

**`ExcelImporter`** - Opens a workbook (`.xlsx`, `.xls`) or CSV file. Dispose it when you're done.
- `ExcelImporter.ReadRows<T>(path or stream, fileType, configure)` - Read every row of the first sheet in one call
- `new ExcelImporter(path or stream)`, `(..., ExcelImporterFileType)`, `(..., ExcelImporterFileType, ExcelReaderConfiguration)` - Open a file, optionally as CSV or with ExcelDataReader options (passwords, CSV separators, fallback encoding); `new ExcelImporter(IExcelDataReader)` wraps a reader you created
- `ReadSheet()` / `ReadSheet(name)` / `ReadSheet(index)` - The next sheet, or a sheet by name (case-insensitive) or zero-based index
- `TryReadSheet(out sheet)` / `TryReadSheet(name, out sheet)` / `TryReadSheet(index, out sheet)` - The same, returning false instead of throwing
- `ReadSheets()` - Every sheet, in order
- `NumberOfSheets` - The number of sheets
- `Configuration` - Settings and class maps (see below)
- `Reader` - The underlying `IExcelDataReader`

**`ExcelSheet`** - A worksheet, and a position in it. Not thread-safe.
- `ReadRows<T>()` - Every row, mapped lazily as you enumerate
- `ReadRows<T>(startIndex, count)` - A block of rows
- `ReadRow<T>()` / `TryReadRow<T>(out value)` - The next row
- `ReadRowResults<T>()` - Every row, with every invalid cell reported in each result instead of thrown ([Collecting errors](errors.md#collecting-errors))
- `ReadHeading()` - Read the heading row now (the read methods do this when needed)
- `FindHeading<T>()` / `FindHeading<T>(maxRows)` / `TryFindHeading<T>(out heading)` - Find the heading row of a sheet with titles, notes or blank rows above it ([Finding the heading](reading.md#finding-the-heading))
- `ValidateHeading<T>()` / `TryValidateHeading<T>(out missingColumns)` - Check every column `T` needs is present, before reading rows
- `Region()` / `Region(ExcelRange)` / `Region("B3:F20")` - A block of the sheet with its own heading ([Reading part of a sheet](reading.md#reading-part-of-a-sheet))
- `HasHeading` (default true), `HeadingIndex` (default 0), `DataRange` (default all rows and columns) - Where the heading and data are
- `Heading` - The heading, once read
- `Name`, `Index`, `Visibility`, `NumberOfColumns`, `CurrentRowIndex`

**`ExcelRegion`** - An immutable block of a sheet, read with its own heading. Each read starts at the top of the region.
- `ReadRows<T>()`, `ReadRowResults<T>()`, `ValidateHeading<T>()`, `TryValidateHeading<T>(out missingColumns)` - As on `ExcelSheet`
- `WithHeadingAt(rowIndex)` / `WithoutHeading()` / `WithHeadingSearch()` / `WithHeadingSearch(maxRows)` - Where the heading is: a row of the sheet, none, or found as for `FindHeading`
- `Sheet`, `Range` - The sheet and block

**`ExcelHeading`** - A heading row.
- `ColumnNames` - The column names, in order
- `GetColumnIndex(name)` / `TryGetColumnIndex(name, out index)` - A column's position in the sheet
- `GetColumnName(index)` - The name of the column at a position
- `GetFirstColumnMatchingIndex(predicate)` / `TryGetFirstColumnMatchingIndex(predicate, out index)` - The first column whose name matches

**`ExcelRange`** - A block of rows and columns, as zero-based ranges.
- `new ExcelRange("A1:B10")`, `ExcelRange.Parse(address)`, `ExcelRange.TryParse(address, out range)` - From an Excel address, including whole rows (`"1:5"`) and columns (`"A:C"`)
- `new ExcelRange(rowStart, rowEnd, columnStart, columnEnd)` - From indexes; `null` ends mean no limit
- `new ExcelRange(rows, columns)` - From `System.Range`s, such as `new ExcelRange(1.., 0..3)` or `new ExcelRange(..^1, Range.All)`
- `Rows`, `Columns` - The ranges
- `ToString()` - An Excel address when the range can be written as one; `default(ExcelRange)` is every row and column

**`ExcelImporterConfiguration`** - An importer's settings and class maps.
- `RegisterClassMap<TMap>()` / `RegisterClassMap(classMap)` / `RegisterClassMap(type, classMap)` - Use a class map for a type. Replaces a map created earlier by auto-mapping; throws if you already registered one
- `RegisterClassMap<T>(c => ...)` - Start from the auto-mapped class map for `T`, and change it
- `TryGetClassMap<T>(out classMap)` / `TryGetClassMap(type, out classMap)` - The class map registered for a type, including one created by auto-mapping
- `Culture` - The culture to parse text and format values with (default: the current culture)
- `SkipBlankLines` - Skip empty rows (default: false)
- `IgnoreSeparatorsInColumnNames` - Match `FirstName` to "First Name", "first_name" and so on (default: false)
- `ValidateDataAnnotations` - Validate each row with data annotations such as `[Range]` (default: false)
- `MaxColumnsPerSheet` - The most columns a sheet may have (default: 10,000)

`importer.RegisterClassMapsInNamespace(namespace)` (and an overload taking an `Assembly`) registers every class map in a namespace.

## Results and errors

**`ExcelRowResult<T>`** - A row read by `ReadRowResults<T>()`.
- `IsValid` - Whether the row mapped without errors
- `Value` - The mapped row (`default` if it isn't valid)
- `Errors` - Every error in the row, each with its cell
- `RowIndex` - The row's zero-based index

**`ExcelMappingException`** - Thrown when a sheet, column or cell can't be mapped.
- `CellAddress` - The cell, such as "B11" (null if unknown)
- `ColumnName` - The column being read, or that couldn't be found (null if unknown)
- `Member` - The property or field being mapped (null if unknown)
- `RowIndex` / `ColumnIndex` - Zero-based row and column (-1 if unknown)
- `Sheet` - The sheet being read
- `InnerException` - Why the value was invalid, for example a `FormatException`

## Class maps

**`ExcelClassMap<T>`** - How rows map to `T`. Derive from it and call `Map(...)` in the constructor, or configure one with `RegisterClassMap<T>(c => ...)`. Members you don't map aren't read.
- `Map(p => p.Member)` - A single value, collection or dictionary member; returns the member's map (below). Indexers map elements one by one: `Map(p => p.Values[0])`, `Map(p => p.Scores["Maths"])`
- `Map(p => p.Status, ignoreCase: true)` - An enum, matched ignoring case
- `MapList<TElement>(p => p.Items)` - A non-generic collection, such as `ArrayList`
- `MapDictionary<TValue>(p => p.Values)` - A dictionary member whose type the `Map` overloads don't cover, with values of type `TValue`
- `MapObject<TElement>(p => p.Address)` - A nested object, mapped from the same row; returns its class map
- `MapRowNumber(p => p.Row)` - A numeric member set to the row's number, as Excel shows it
- `MapExtraColumns(p => p.Other)` - A dictionary of the columns no other member reads
- `WithClassMap(classMap)` / `WithClassMap(c => ...)` - On a nested object's class map (from `MapObject`), use another class map or configure this one: `MapObject(p => p.Address).WithClassMap(c => c.Map(a => a.City))`
- `new ExcelClassMap<T>(FallbackStrategy.SetToDefaultValue)` - Give empty cells the type's default value instead of an error

### Single values (`OneToOneMap<T>`)

Columns:
- `.WithColumnName("Name")` / `.WithColumnName("Name", StringComparison.Ordinal)` - A column by name (ignoring case by default)
- `.WithColumnNameMatching("Name1", "Name2")` - The first of these columns that the sheet has
- `.WithColumnNameMatching(name => name.StartsWith("Total"))` / `.WithColumnNameMatching(regex)` - The first column whose name matches
- `.WithColumnIndex(0)` / `.WithColumnLetter("C")` - A column by position
- `.WithColumns(selector)` - The first column an `IColumnSelector` chooses
- `.MakeOptional()` - Leave the member unset if the column is missing
- `.MakePreserveFormatting()` - Read the cell's text as Excel displays it

Values (also available on collection elements and dictionary values, through `WithElementMap` and `WithValueMap`):
- `.WithConverter(value => ...)` - Convert the cell's text with a delegate, replacing the built-in conversion
- `.WithConverter(ICellConverter<T>)` - Convert with a converter, which gets the typed cell value and where it is ([Custom converters](mapping.md#custom-converters))
- `.WithMapping(dictionary, comparer, behavior)` - Map fixed text to values
- `.WithFormats("yyyy-MM-dd", ...)` - Date, time and duration formats
- `.WithFormatProvider(provider)` - The culture for dates, times, durations and (.NET 8) numbers
- `.WithNumberStyle(NumberStyles.Currency)` - How numbers are parsed (.NET 8)
- `.WithOffset(TimeSpan)` - The offset of `DateTimeOffset` values (dates in Excel have no time zone)
- `.WithUriKind(UriKind.Relative)` - Which URIs are valid (absolute only by default)
- `.WithTrim()` / `.WithTransformers(ICellTransformer...)` - Change cells before they're converted

Fallbacks:
- `.WithEmptyFallback(value)` / `.WithInvalidFallback(value)` / `.WithValueFallback(value)` - The value of empty cells, invalid cells, or both
- `.WithEmptyFallback(ICellFallback<T>)` / `.WithInvalidFallback(ICellFallback<T>)` - Compute the value ([Fallbacks](mapping.md#advanced-fallback-strategies))
- `.WithThrowingEmptyFallback()` / `.WithThrowingInvalidFallback()` / `.WithThrowingFallback()` - Throw instead (the default for non-nullable value types)

For a nullable member (`int?`), converters and fallbacks can be for the underlying type: `WithConverter(ICellConverter<int>)`, `WithEmptyFallback(ICellFallback<int>)`.

### Collections (`ManyToOneEnumerableMap<TElement>`) and dictionaries (`ManyToOneDictionaryMap<TKey, TValue>`)

- `.WithColumnNames("A", "B")` / `.WithColumnIndices(0, 1)` - Read these columns
- `.WithColumnsMatching(predicate)` / `.WithColumnsMatching(regex)` - Read every matching column
- `.WithColumns(selector)` - Read the columns an `IColumnSelector` chooses
- `.WithColumnName("Tags")` / `.WithColumnIndex(0)` with `.WithSeparators(',', ';')` - Split one cell (collections only; `,` by default)
- `.WithElementMap(e => ...)` / `.WithValueMap(v => ...)` - Configure each element or value with the value methods above
- `.MakeOptional()`, `.MakePreserveFormatting()` - As for single values

## Attributes

| Attribute | Purpose |
|-----------|---------|
| `[ExcelColumnName("Name")]` | A column by name (ignoring case; pass a `StringComparison` to change that) |
| `[ExcelColumnNames("Name1", "Name2")]` | For a single value, the first of these columns the sheet has; for a collection or dictionary, all of them |
| `[ExcelColumnIndex(0)]` / `[ExcelColumnLetter("C")]` | A column by position |
| `[ExcelColumnIndices(0, 1)]` | For a single value, the first of these columns the sheet has; for a collection or dictionary, all of them |
| `[ExcelColumnMatching("regex")]` / `[ExcelColumnMatching(typeof(Selector))]` | The first column matching a pattern, or chosen by an `IColumnSelector` |
| `[ExcelColumnsMatching("regex")]` / `[ExcelColumnsMatching(typeof(Selector))]` | Every column matching a pattern, or chosen by an `IColumnSelector` (collections and dictionaries) |
| `[ExcelSeparators(',', ';')]` | Split one cell into a collection with these separators |
| `[ExcelOptional]` | Leave the member unset if the column is missing |
| `[ExcelIgnore]` | Don't map the member |
| `[ExcelRowNumber]` | Set a numeric member to the row's number |
| `[ExcelExtraColumns]` | Fill a dictionary with the columns no other member reads |
| `[ExcelPreserveFormatting]` | Read the cell's text as Excel displays it |
| `[ExcelTrimString]` | Trim text before it's converted |
| `[ExcelTransformer(typeof(Transformer))]` | Change cells with an `ICellTransformer` |
| `[ExcelConverter(typeof(Converter))]` | Convert with an `ICellConverter<T>`, replacing the built-in conversion |
| `[ExcelDefaultValue(value)]` / `[ExcelInvalidValue(value)]` | The value of empty or invalid cells |
| `[ExcelEmptyFallback(typeof(Fallback))]` / `[ExcelInvalidFallback(typeof(Fallback))]` | Compute the value of empty or invalid cells with an `ICellFallback<T>` |
| `[ExcelFormats("yyyy-MM-dd", ...)]` | Date, time and duration formats |
| `[ExcelNumberStyle(NumberStyles.Currency)]` | How numbers are parsed |
| `[ExcelOffset("+01:00")]` | The offset of `DateTimeOffset` values |
| `[ExcelUri(UriKind.Relative)]` | Which URIs are valid: `Absolute` (the default), `Relative` or `RelativeOrAbsolute` |
| `[ExcelMappingDictionary("text", value)]` | Map fixed text to a value (repeat for each value) |
| `[ExcelMappingDictionaryComparer(StringComparison.OrdinalIgnoreCase)]` | How the text is compared |
| `[ExcelMappingDictionaryBehavior(MappingDictionaryMapperBehavior.Required)]` | Whether text that isn't in the dictionary is invalid |

Attributes that take a type (`[ExcelConverter]`, `[ExcelTransformer]`, the fallbacks and the column selectors) also take `ConstructorArguments`.

## Extension points

| Interface | Implement to | Set with |
|---|---|---|
| `ICellConverter<T>` | Convert a cell (`CellValue`, with `CellContext`) to `T`, returning a `CellConversionResult<T>`: `Success`, `Invalid` or `NotHandled` | `WithConverter(converter)`, `[ExcelConverter]` |
| `ICellTransformer` | Change a cell's value before it's converted | `WithTransformers(...)`, `[ExcelTransformer]` |
| `ICellFallback<T>` | Give the value of an empty or invalid cell | `WithEmptyFallback`/`WithInvalidFallback(fallback)`, `[ExcelEmptyFallback]`/`[ExcelInvalidFallback]` |
| `IColumnSelector` | Choose the columns a member is read from, once for each sheet | `WithColumns(selector)`, `[ExcelColumnMatching(typeof(...))]`, `[ExcelColumnsMatching(typeof(...))]` |

**`CellValue`** - A cell's value: `Kind` (`Empty`, `Text`, `Number`, `Boolean`, `DateTime`, `TimeSpan`, `Other`), `IsEmpty`, `Text`, `Value` (as Excel stores it), and `Number`, `Boolean`, `DateTime` and `TimeSpan` for typed cells. Create one with `new CellValue(text or number or bool or DateTime or TimeSpan)`, for example to return from a transformer.

**`CellContext`** - Where a cell is: `Sheet`, `RowIndex`, `ColumnIndex`, `CellAddress`, `Member`, and `FormatProvider` (the importer's culture).
