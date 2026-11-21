# API Reference

The main classes, attributes and fluent methods. See the [documentation index](../README.md#documentation) for other topics.

## API Reference

### Core Classes

**`ExcelImporter`** - Main entry point for reading Excel files
- `ExcelImporter.ReadRows<T>(path or stream, fileType, configure)` - Read every row of the first sheet in one call
- `ReadSheet()` / `TryReadSheet()` - Read sheets sequentially or by name/index
- `ReadSheets()` - Enumerate all sheets
- `Configuration` - Access configuration for registering maps
- `Reader` - Access underlying `IExcelDataReader` for advanced scenarios
- `NumberOfSheets` - Get total sheet count

**`ExcelSheet`** - Represents a single worksheet
- `ReadRows<T>()` - Read all rows as typed objects (lazy evaluation)
- `ReadRows<T>(startIndex, count)` - Read specific range of rows
- `ReadRow<T>()` / `TryReadRow<T>()` - Read single row
- `ReadRowResults<T>()` - Read all rows, reporting every invalid cell instead of throwing
- `ReadHeading()` - Explicitly read header row
- `ValidateHeading<T>()` / `TryValidateHeading<T>()` - Check that every column a type needs is present, before reading rows
- `Region()` / `Region(range)` / `Region("B3:F20")` - A block of the sheet with its own heading, to read part of a sheet or several tables on one sheet

**`ExcelRegion`** - A block of a sheet, read with its own heading (immutable)
- `ReadRows<T>()`, `ReadRowResults<T>()`, `ValidateHeading<T>()`, `TryValidateHeading<T>()` - As on `ExcelSheet`, reading from the top of the region each time
- `WithHeadingAt(rowIndex)` / `WithoutHeading()` / `WithHeadingSearch()` - Where the heading is
- `Sheet`, `Range` - The sheet and block
- `FindHeading<T>()` / `TryFindHeading<T>()` - Find and read the heading row when there are titles or notes above it
- `Name` - Sheet name
- `Visibility` - Sheet visibility (Visible, Hidden, VeryHidden)
- `Index` - Zero-based sheet index
- `NumberOfColumns` - Column count
- `HasHeading` - Whether sheet has header row (default: true)
- `HeadingIndex` - Zero-based index of header row (default: 0)
- `DataRange` - Range of rows and columns to process (default: all)
- `CurrentRowIndex` - Current row being processed

**`ExcelRange`** - Represents a range of rows and columns
- Constructor: `new ExcelRange(rowStart, rowEnd, columnStart, columnEnd)` - Create range with explicit indices
- Constructor: `new ExcelRange(rows, columns)` - Create range with Range objects
- Constructor: `new ExcelRange(address)` - Create range from Excel address string (e.g., "A1:B10")
- `Parse(string)` - Parse Excel address string (throws on error)
- `TryParse(string, out ExcelRange)` - Try parse Excel address string (returns bool)
- `Rows` - Row range (System.Range)
- `Columns` - Column range (System.Range)
- `ToString()` - Excel address such as "A1:B10", "A:C" or "1:5" when the range can be written as one
- Equality: `Equals`, `==` and `!=` compare rows and columns; `default(ExcelRange)` means all rows and columns

**`ExcelMappingException`** - Thrown when a sheet or cell can't be mapped
- `CellAddress` - Excel address of the cell, such as "B11" (null if unknown)
- `ColumnName` - Name of the column being read, or of a column that couldn't be found (null if unknown)
- `Member` - The property or field being mapped (null if unknown)
- `RowIndex` / `ColumnIndex` - Zero-based row and column of the cell (-1 if unknown)
- `Sheet` - The sheet being read

**`ExcelImporterConfiguration`** - Configuration settings
- `RegisterClassMap<T>()` - Register type-specific mapping. Replaces a map created earlier by auto-mapping; throws if you already registered a map for the type
- `RegisterClassMap<T>(Action<ExcelClassMap<T>>)` - Inline lambda configuration
- `TryGetClassMap<T>()` - Check if map exists
- `SkipBlankLines` - Skip empty rows (default: false)
- `IgnoreSeparatorsInColumnNames` - Ignore spaces, underscores, hyphens and dots when matching column names (default: false)
- `Culture` - Culture used to parse text and convert numbers and dates to text (default: null, the current culture)
- `MaxColumnsPerSheet` - Security limit (default: 10,000)
- `ValidateDataAnnotations` - Enable data annotations validation (default: false)

**`ExcelClassMap<T>`** - Fluent mapping configuration
- `Map(expression)` - Map property or field
- `MapObject<TElement>(expression)` - Map nested object
- `MapRowNumber(expression)` - Set a numeric member to the row number
- `MapExtraColumns(expression)` - Fill a dictionary member with the columns no other member reads
- `MapEnumerable<TElement>(expression)` - Map collection
- `MapDictionary<TKey, TValue>(expression)` - Map dictionary

### Mapping Attributes

| Attribute | Purpose |
|-----------|---------|
| `[ExcelColumnName("Name")]` | Map to specific column name (case-insensitive by default) |
| `[ExcelColumnName("Name", StringComparison.Ordinal)]` | Map to column name with specific comparison mode |
| `[ExcelColumnNames("Name1", "Name2")]` | Try multiple column names |
| `[ExcelColumnIndex(0)]` | Map to column by index |
| `[ExcelColumnLetter("C")]` | Map to column by Excel letter |
| `[ExcelColumnIndices(0, 1)]` | Try multiple column indices |
| `[ExcelColumnMatching(@"regex", options)]` | Match columns by pattern |
| `[ExcelColumnsMatching(typeof(Selector))]` | Custom column matching (`IColumnSelector`) |
| `[ExcelOptional]` | Don't throw if column missing |
| `[ExcelIgnore]` | Exclude from mapping |
| `[ExcelRowNumber]` | Set to the one-based row number instead of reading a column |
| `[ExcelExtraColumns]` | Fill a dictionary with the columns no other member reads |
| `[ExcelDefaultValue(value)]` | Default for empty cells |
| `[ExcelInvalidValue(value)]` | Default for invalid values |
| `[ExcelEmptyFallback(typeof(Fallback))]` | Custom empty cell handling (`ICellFallback<T>`) |
| `[ExcelInvalidFallback(typeof(Fallback))]` | Custom invalid value handling (`ICellFallback<T>`) |
| `[ExcelPreserveFormatting]` | Read formatted string |
| `[ExcelTrimString]` | Auto-trim whitespace |
| `[ExcelFormats("format1", "format2")]` | Parse dates/times with specific formats |
| `[ExcelNumberStyle(NumberStyles.AllowThousands)]` | Control numeric parsing (thousands, currency, etc.) |
| `[ExcelOffset("+00:00")]` | Offset from UTC for `DateTimeOffset` members |
| `[ExcelUri(UriKind.RelativeOrAbsolute)]` | Specify URI kind (Absolute, Relative, or RelativeOrAbsolute) |
| `[ExcelSeparators(';', ',')]` | Split cell value with custom separators |
| `[ExcelTransformer(typeof(Transformer))]` | Apply custom transformer |
| `[ExcelConverter(typeof(Converter))]` | Convert with an `ICellConverter<T>`, replacing the built-in conversion |
| `[ExcelMappingDictionary("key", value)]` | Map string value to enum/object (multiple allowed) |
| `[ExcelMappingDictionaryComparer(comparison)]` | Set string comparison for dictionary keys |
| `[ExcelMappingDictionaryBehavior(behavior)]` | Control required vs optional mapping |

### Fluent API Methods

**Column Selection:**
- `.WithColumnName("Name")` - Map to specific column
- `.WithColumnNames("Name1", "Name2")` - Try multiple names
- `.WithColumnIndex(0)` - Map by index
- `.WithColumnLetter("C")` - Map by Excel column letter
- `.WithColumnIndices(0, 1, 2)` - Try multiple indices
- `.WithColumnNameMatching(predicate)` / `.WithColumnNameMatching(regex)` - Use predicate or regular expression
- `.WithColumnsMatching(predicate | regex)` - All matching columns (collections and dictionaries)
- `.WithColumns(IColumnSelector)` - Columns chosen by a selector

**Behavior:**
- `.MakeOptional()` - Don't throw if missing
- `.WithEmptyFallback(value)` - Default for empty
- `.WithInvalidFallback(value)` - Default for invalid
- `.WithValueFallback(value)` - Default for both
- `.WithConverter(value => ...)` - Custom conversion, replacing the built-in conversion
- `.WithConverter(ICellConverter<T>)` - Custom conversion with the typed cell value and its context
- `.WithFormats("format1", "format2")` - Date/time formats
- `.WithFormatProvider(IFormatProvider)` - Culture-specific parsing for dates/times/numbers
- `.WithNumberStyle(NumberStyles)` - Control numeric parsing (thousands, currency, etc.)
- `.WithUriKind(UriKind)` - Specify URI kind for Uri properties
- `.WithOffset(TimeSpan)` - Offset from UTC for `DateTimeOffset` properties (dates in Excel have no time zone)
- `.WithMapping(dictionary)` - Value mapping
- `.WithTrim()` - Trim whitespace
- `.WithTransformers(...)` - Custom transformers (`ICellTransformer`)
- `.WithEmptyFallback(ICellFallback<T>)` / `.WithInvalidFallback(ICellFallback<T>)` - Compute the value of empty or invalid cells
- `.WithMappers(...)` - Custom mappers

**Collections:**
- `.WithSeparators(';', ',')` - Split delimiters (char)
- `.WithSeparators(";", ",")` - Split delimiters (string)
- `.WithElementMap(m => ...)` - Configure element pipeline

### Extension Interfaces

Implement these for advanced customization:

- `ICellConverter<T>` - Convert a cell's value (`CellValue`, with `CellContext`) to `T`, returning a `CellConversionResult<T>`
- `ICellTransformer` (in the `ExcelMapper` namespace) - Change a cell's value (`CellValue`) before it's converted
- `ICellFallback<T>` - Give the value of an empty or invalid cell
- `IColumnSelector` - Choose the columns a member is read from, once for each sheet
- `ICellMapper` - Custom type conversion logic
- `ICellTransformer` - Transform string values before mapping
- `IFallbackItem` - Custom fallback behavior
