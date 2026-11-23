# API Reference

The public types, attributes and fluent methods, grouped by what they do. Every type is in the `ExcelMapper` namespace. The package includes XML documentation, so each member's details also show in your editor. See the [documentation index](../README.md#documentation) for guides.

## Reading files

**`ExcelImporter`** - Opens a workbook (`.xlsx`, `.xls`, `.xlsb`) or CSV file. Dispose it when you're done.
- `ExcelImporter.ReadRows<T>(path or stream)`, optionally with a file type and a `configure` callback - Read every row of the first sheet in one call. A path ending in `.csv`, `.tsv` or `.txt` is read as CSV
- `new ExcelImporter(path or stream)` (a path ending in `.csv`, `.tsv` or `.txt` is CSV), `(..., ExcelImporterFileType)`, `(..., ExcelImporterFileType, ExcelReaderConfiguration)` - Open a file, optionally as CSV or with ExcelDataReader options (passwords, CSV separators, fallback encoding); `new ExcelImporter(IExcelDataReader)` wraps a reader you created
- `ReadSheet()` / `ReadSheet(name)` / `ReadSheet(index)` - The next sheet, or a sheet by name (case-insensitive) or zero-based index
- `TryReadSheet(out sheet)` / `TryReadSheet(name, out sheet)` / `TryReadSheet(index, out sheet)` - The same, returning false instead of throwing
- `ReadSheets()` - Every sheet, in order
- `NumberOfSheets` - The number of sheets
- `Configuration` - Settings and class maps (see below)

**`ExcelSheet`** - A worksheet, and a position in it. Not thread-safe.
- `ReadRows<T>()` - Every row, mapped lazily as you enumerate
- `ReadRows<T>(startIndex, count)` - A block of rows
- `ReadRow<T>()` / `TryReadRow<T>(out value)` - The next row
- `ReadRowResults<T>()` - Every row, with every invalid cell reported in each result instead of thrown ([Collecting errors](errors.md#collecting-errors))
- `ReadRows<T>(IExcelErrorSink errors)` - The rows without problems, giving each problem to `errors` without creating an exception for it ([Checking large files](errors.md#checking-large-files))
- `ReadHeading()` - Read the heading row now (the read methods do this when needed)
- `FindHeading<T>()` / `FindHeading<T>(maxRows)` / `TryFindHeading<T>(out heading)` - Find the heading row of a sheet with titles, notes or blank rows above it ([Finding the heading](reading.md#finding-the-heading))
- `ValidateHeading<T>()` / `TryValidateHeading<T>(out missingColumns)` - Check every column `T` needs is present, before reading rows
- Each of these also takes an `ExcelClassMap<T>`, such as `ReadRows(classMap)`: it reads with that class map rather than the one registered or auto-mapped for `T`, and works with native AOT ([Native AOT apps](trimming.md#native-aot-apps))
- `Region()` / `Region(ExcelRange)` / `Region("B3:F20")` - A block of the sheet with its own heading ([Reading part of a sheet](reading.md#reading-part-of-a-sheet))
- `HasHeading` (default true), `HeadingIndex` (default 0), `DataRange` (default all rows and columns) - Where the heading and data are
- `Heading` - The heading, once read
- `Name`, `Index`, `Visibility`, `NumberOfColumns`, `CurrentRowIndex`

**`ExcelRegion`** - An immutable block of a sheet, read with its own heading. Each read starts at the top of the region.
- `ReadRows<T>()`, `ReadRows<T>(errors)`, `ReadRowResults<T>()`, `ValidateHeading<T>()`, `TryValidateHeading<T>(out missingColumns)` - As on `ExcelSheet`, also with an `ExcelClassMap<T>`
- `WithHeadingAt(rowIndex)` / `WithoutHeading()` / `WithHeadingSearch()` / `WithHeadingSearch(maxRows)` - Where the heading is: a row of the sheet, none, or found as for `FindHeading`
- `Sheet`, `Range` - The sheet and block

**`ExcelHeading`** - A heading row.
- `ColumnNames` - The column names, in order
- `GetColumnIndex(name)` / `TryGetColumnIndex(name, out index)` - A column's position in the sheet
- `GetColumnName(index)` - The name of the column at a position
- `GetFirstColumnMatchingIndex(predicate)` / `TryGetFirstColumnMatchingIndex(predicate, out index)` - The first column whose name matches

**`ExcelRange`** - A block of rows and columns, as zero-based ranges.
- `ExcelRange.Parse("A1:B10")`, `ExcelRange.TryParse(address, out range)` - From an Excel address, including whole rows (`"1:5"`) and columns (`"A:C"`)
- `new ExcelRange(rowStart, rowEnd, columnStart, columnEnd)` - From indexes; `null` ends mean no limit
- `new ExcelRange(rows, columns)` - From `System.Range`s, such as `new ExcelRange(1.., 0..3)` or `new ExcelRange(..^1, Range.All)`
- `Rows`, `Columns` - The ranges
- `ToString()` - An Excel address when the range can be written as one; `default(ExcelRange)` is every row and column

**`ExcelImporterConfiguration`** - An importer's settings and class maps.
- `RegisterClassMap<TMap>()` / `RegisterClassMap(classMap)` - Use a class map for its type. Replaces a map created earlier by auto-mapping; throws if you already registered one
- `RegisterClassMap<T>(c => ...)` - Configure a class map for `T` inline. It starts empty: only the members you map are read, unless you call `c.AutoMap()`
- `TryGetClassMap<T>(out classMap)` (an `ExcelClassMap<T>`) / `TryGetClassMap(type, out classMap)` - The class map registered for a type, including one created by auto-mapping
- `Culture` - The culture to parse text and format values with (default: `CultureInfo.InvariantCulture`)
- `SkipBlankLines` - Skip empty rows (default: false)
- `IgnoreSeparatorsInColumnNames` - Match `FirstName` to "First Name", "first_name" and so on (default: false)
- `ValidateDataAnnotations` - Validate each row with data annotations such as `[Range]` (default: false)
- `MaxColumnsPerSheet` - The most columns a sheet may have (default: 10,000)

`importer.Configuration.RegisterClassMapsInNamespace(namespace)` (and an overload taking an `Assembly`) registers every class map in a namespace.

**`ExcelMappingContext`** - Class maps for native AOT ([Native AOT apps](trimming.md#native-aot-apps)). Declare a `partial` class deriving from it with `[ExcelMappable(typeof(T))]` for each row type, and the source generator writes it; or register types in `RegisterTypes(ExcelTypeRegistrar)` yourself.
- `GetClassMap<T>()` - The class map of a registered row type, auto-mapped from the registered types; it can't be changed

**`ExcelMappableAttribute`** - `[ExcelMappable(typeof(T))]` on a `partial` context: the source generator registers `T` and every type its members use, and adds a property with `T`'s class map.

**`ExcelTypeRegistrar`** - Registers types in `RegisterTypes`, for a context written by hand: `Row<T>()`, `Type<T>()`, `Nullable<T>()`, `Enum<TEnum>()`, `Parsable<T>()`, `ParsableClass<T>()`, `Collection<TCollection, TElement>()`, `Array<TElement>()` and `Dictionary<TDictionary, TKey, TValue>()`. `Setter` and `Constructor` register code that sets a member (`ExcelMemberSetter<TOwner, TValue>`) or calls a constructor (`ExcelConstructor<TOwner>`, with `ExcelConstructorArguments<TOwner>`), which the generator writes so native AOT apps don't use reflection.

## Results and errors

**`IExcelErrorSink`** - Receives the problems found by `ReadRows<T>(IExcelErrorSink)`: implement `Add(in ExcelMappingError error)`.

**`ExcelMappingError`** - A problem given to an `IExcelErrorSink`, only valid during `Add`: `Kind` (`ExcelMappingErrorKind.EmptyValue`, `InvalidValue`, `DataAnnotation` or `Exception`), `Sheet`, `RowIndex`, `ColumnIndex`, `ColumnName`, `Member`, the cell's `Value`, and `ToException()`.

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
- `Map(p => p.Status).WithIgnoreCase()` - An enum, matched ignoring case (also on an enum collection's `WithElementMap`)
- `MapCollection<TElement>(p => p.Items)` - A collection whose type the `Map` overloads don't cover, such as `HashSet<T>` or `ArrayList`
- `MapDictionary<TValue>(p => p.Values)` - A dictionary member whose type the `Map` overloads don't cover, with values of type `TValue`
- `MapObject<TElement>(p => p.Address)` - A nested object, mapped from the same row; returns its class map
- `AutoMap()` - Map every member auto-mapping would, from its type and attributes, except members already mapped; members mapped afterwards replace the auto-mapped ones
- `MapRowNumber(p => p.Row)` - A numeric member set to the row's number, as Excel shows it
- `MapExtraColumns(p => p.Other)` - A dictionary of the columns no other member reads
- `WithClassMap(classMap)` / `WithClassMap(c => ...)` - On a nested object's class map (from `MapObject`), use another class map or configure this one: `MapObject(p => p.Address).WithClassMap(c => c.Map(a => a.City))`
- `new ExcelClassMap<T>(EmptyValueStrategy.DefaultValue)` - Give empty cells the type's default value instead of an error

### Single values (`ValueMap<T>`)

Columns:
- `.WithColumnName("Name")` / `.WithColumnName("Name", StringComparison.Ordinal)` - A column by name (ignoring case by default)
- `.WithColumnNames("Name1", "Name2")` - The first of these columns that the sheet has
- `.WithColumnNameMatching(name => name.StartsWith("Total"))` / `.WithColumnNameMatching(regex)` - The first column whose name matches
- `.WithColumnIndex(0)` / `.WithColumnLetter("C")` - A column by position
- `.WithColumns(selector)` - The first column an `IColumnSelector` chooses
- `.MakeOptional()` - Leave the member unset if the column is missing
- `.MakePreserveFormatting()` - Read the cell's text as Excel displays it

Values (also available on collection elements and dictionary values, through `WithElementMap` and `WithValueMap`):
- `.WithConverter(value => ...)` - Convert the cell's text with a delegate, replacing the built-in conversion
- `.WithConverter(ICellConverter<T>)` - Convert with a converter, which gets the typed cell value and where it is ([Custom converters](mapping.md#converters))
- `.WithMapping(dictionary)`, optionally with a `comparer` and a `MappingDictionaryBehavior` - Map fixed text to values
- `.WithFormats("yyyy-MM-dd", ...)` - Date, time and duration formats
- `.WithFormatProvider(provider)` - The culture for dates, times, durations and numbers
- `.WithNumberStyle(NumberStyles.Currency)` - How numbers are parsed (on the netstandard builds, for the built-in number types)
- `.WithOffset(TimeSpan)` - The offset of `DateTimeOffset` values without one (UTC by default)
- `.WithUriKind(UriKind.Relative)` - Which URIs are valid (absolute only by default)
- `.WithTrim()` / `.WithTransformers(ICellTransformer...)` - Change cells before they're converted

Fallbacks:
- `.WithEmptyFallback(value)` / `.WithInvalidFallback(value)` / `.WithFallback(value)` - The value of empty cells, invalid cells, or both
- `.WithEmptyFallback(ICellFallback<T>)` / `.WithInvalidFallback(ICellFallback<T>)` - Compute the value ([Fallbacks](mapping.md#computed-fallbacks))
- `.WithThrowingEmptyFallback()` / `.WithThrowingInvalidFallback()` / `.WithThrowingFallback()` - Throw instead (the default for non-nullable value types)

For a nullable member (`int?`), converters and fallbacks can be for the underlying type: `WithConverter(ICellConverter<int>)`, `WithEmptyFallback(ICellFallback<int>)`.

### Collections (`CollectionMap<TElement>`) and dictionaries (`DictionaryMap<TKey, TValue>`)

- `.WithColumnNames("A", "B")` / `.WithColumnIndices(0, 1)` - Read these columns
- `.WithColumnsMatching(predicate)` / `.WithColumnsMatching(regex)` - Read every matching column
- `.WithColumns(selector)` - Read the columns an `IColumnSelector` chooses
- `.WithColumnName("Tags")` / `.WithColumnIndex(0)` / `.WithColumnLetter("B")` with `.WithSeparators(',', ';')` and `.WithSplitOptions(StringSplitOptions.RemoveEmptyEntries)` - Split one cell (collections only; `,` by default)
- `.WithElementMap(e => ...)` / `.WithValueMap(v => ...)` - Configure each element or value with the value methods above
- `.MakeOptional()`, `.MakePreserveFormatting()` - As for single values

## Attributes

| Attribute | Purpose |
|-----------|---------|
| `[ExcelColumnName("Name")]` | A column by name (ignoring case; set `Comparison` to change that) |
| `[ExcelColumnNames("Name1", "Name2")]` | For a single value, the first of these columns the sheet has; for a collection or dictionary, all of them (set `Comparison` to change how names are compared) |
| `[ExcelColumnIndex(0)]` / `[ExcelColumnLetter("C")]` | A column by position |
| `[ExcelColumnIndices(0, 1)]` | For a single value, the first of these columns the sheet has; for a collection or dictionary, all of them |
| `[ExcelColumnMatching("regex")]` / `[ExcelColumnMatching(typeof(Selector))]` | The first column matching a pattern (set `RegexOptions` for its options), or chosen by an `IColumnSelector` |
| `[ExcelColumnsMatching("regex")]` / `[ExcelColumnsMatching(typeof(Selector))]` | Every column matching a pattern (set `RegexOptions` for its options), or chosen by an `IColumnSelector` (collections and dictionaries) |
| `[ExcelSeparators(',', ';')]` | Split one cell into a collection with these separators (set `Options` to remove empty or trim values) |
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
| `[ExcelOffset("+01:00")]` | The offset of `DateTimeOffset` values without one (UTC by default) |
| `[ExcelUriKind(UriKind.Relative)]` | Which URIs are valid: `Absolute` (the default), `Relative` or `RelativeOrAbsolute` |
| `[ExcelMappingDictionary("text", value)]` | Map fixed text to a value (repeat for each value) |
| `[ExcelMappingDictionaryComparer(StringComparison.OrdinalIgnoreCase)]` | How the text is compared |
| `[ExcelMappingDictionaryBehavior(MappingDictionaryBehavior.Required)]` | Whether text that isn't in the dictionary is invalid |

Attributes that take a type (`[ExcelConverter]`, `[ExcelTransformer]`, the fallbacks and the column selectors) also take `ConstructorArguments`.

## Extension points

| Interface | Implement to | Set with |
|---|---|---|
| `ICellConverter<T>` | Convert a cell (`CellValue`, with `CellContext`) to `T`, returning a `CellConversionResult<T>`: `Success`, `Invalid` or `NotHandled` | `WithConverter(converter)`, `[ExcelConverter]` |
| `ICellTransformer` | Change a cell's value before it's converted | `WithTransformers(...)`, `[ExcelTransformer]` |
| `ICellFallback<T>` | Give the value of an empty or invalid cell | `WithEmptyFallback`/`WithInvalidFallback(fallback)`, `[ExcelEmptyFallback]`/`[ExcelInvalidFallback]` |
| `IColumnSelector` | Choose the columns a member is read from, once for each sheet, from a `ColumnSelectorContext` | `WithColumns(selector)`, `[ExcelColumnMatching(typeof(...))]`, `[ExcelColumnsMatching(typeof(...))]` |

**`CellValue`** - A cell's value: `Kind` (`Empty`, `Text`, `Number`, `Boolean`, `DateTime`, `TimeSpan`, `Other`), `IsEmpty`, `Text`, `Value` (as Excel stores it), and `Number`, `Boolean`, `DateTime` and `TimeSpan` for typed cells. Create one with `new CellValue(text or number or bool or DateTime or TimeSpan)`, for example to return from a transformer.

**`CellContext`** - Where a cell is: `Sheet`, `Heading` (a region's, when reading one), `RowIndex`, `ColumnIndex`, `CellAddress`, `Member`, and `FormatProvider` (the importer's culture). `new ExcelMappingException(message, context)` reports the cell as an error.

**`ColumnSelectorContext`** - What an `IColumnSelector` chooses from: `Sheet`, `Heading` (a region's, when reading one) and `NumberOfColumns`.
