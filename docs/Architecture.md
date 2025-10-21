# ExcelMapper Architecture

## Overview

ExcelMapper is a .NET library for mapping Excel spreadsheet data to strongly-typed objects. It provides a flexible, extensible pipeline-based architecture that supports various mapping scenarios with robust type safety and error handling.

## Core Concepts

### Mapping Types

ExcelMapper supports four primary mapping patterns:

1. **One-to-One Mapping**: A single cell value maps to a single property/field
   - Example: Cell "A1" → `Person.Name`

2. **One-to-Many Mapping**: A single cell value maps to multiple values (typically via splitting)
   - Example: Cell "A1" containing "John,Jane,Bob" → `string[] Names`

3. **Many-to-One Mapping**: Multiple cell values map to a single collection/dictionary property
   - Example: Cells "A1", "B1", "C1" → `List<string> Values`

4. **Many-to-Many Mapping**: Multiple cell values map to multiple values in a collection
   - Example: Multiple cells with comma-separated values → `List<List<string>>`

### Map Hierarchy

The library uses a hierarchical map structure:

```
IMap (base interface)
├── IToOneMap (single property/field)
│   ├── IOneToOneMap → OneToOneMap<T>
│   └── IManyToOneMap → ManyToOneEnumerableMap<T>
│                     → ManyToOneDictionaryMap<TKey, TValue>
└── Indexer Maps (collection element access)
    ├── IEnumerableIndexerMap → ManyToOneEnumerableIndexerMap<T>
    ├── IDictionaryIndexerMap → ManyToOneDictionaryIndexerMap<TKey, TValue>
    └── IMultidimensionalIndexerMap → ManyToOneMultidimensionalIndexerMap<T>
```

## Architecture Layers

### 1. Import Layer

**Entry Point**: `ExcelImporter`

The importer is responsible for reading Excel files and providing access to sheets:

```csharp
using var importer = new ExcelImporter("data.xlsx");
ExcelSheet sheet = importer.ReadSheet();
var records = sheet.ReadRows<Person>();
```

**Key Components**:
- **`ExcelImporter`**: Main entry point, wraps `IExcelDataReader` from ExcelDataReader library
  - Supports Excel (.xlsx, .xls) and CSV files
  - Manages sheet navigation and reader state
  - Provides configuration via `ExcelImporterConfiguration`

- **`ExcelSheet`**: Represents a single worksheet
  - Manages heading/header row detection
  - Provides row enumeration with `ReadRows<T>()`
  - Supports range-based reading with `ReadRows<T>(startIndex, count)`
  - Handles blank line skipping

- **`ExcelHeading`**: Represents the header row
  - Maps column names to zero-based indices
  - Enables name-based column lookup

### 2. Configuration Layer

**Entry Point**: `ExcelImporterConfiguration`

Configuration controls mapping behavior and stores registered class maps:

**Key Settings**:
- **`SkipBlankLines`**: Whether to skip empty rows (default: false)
- **`RegisteredMaps`**: Dictionary of registered `ExcelClassMap` instances per type

**Class Maps**:
- **`ExcelClassMap<T>`**: Defines how to map rows to objects of type `T`
  - Created manually via `new ExcelClassMap<Person>()`
  - Created automatically via AutoMapper for unmapped types
  - Contains collection of `ExcelPropertyMap` instances

### 3. Mapping Layer

#### Property Maps

Each property/field has a corresponding map that defines how to read and convert data:

**`OneToOneMap<T>`**: Maps a single cell to a property
- Contains `ICellReaderFactory` to create readers
- Contains `ValuePipeline<T>` for transformation and conversion
- Supports `Optional` flag for missing columns
- Supports `PreserveFormatting` to retain Excel formatting

**`ManyToOneEnumerableMap<TElement>`**: Maps multiple cells to a collection
- Uses `ICellsReaderFactory` for multi-cell reading
- Uses `IEnumerableFactory<TElement>` to construct the target collection type
- Contains `IValuePipeline<TElement>` for element conversion
- Supports List, Array, HashSet, ImmutableList, and custom collection types

**`ManyToOneDictionaryMap<TKey, TValue>`**: Maps multiple cells to a dictionary
- Keys are derived from column names
- Values are mapped using `IValuePipeline<TValue>`
- Uses `IDictionaryFactory<TKey, TValue>` to construct dictionaries
- Supports Dictionary, FrozenDictionary, ImmutableDictionary

#### Indexer Maps

Used when mapping to collection elements via indexer syntax (e.g., `list[0]`, `dict["key"]`):

- **`ManyToOneEnumerableIndexerMap<T>`**: For `list[0]` style access
- **`ManyToOneDictionaryIndexerMap<TKey, TValue>`**: For `dict["key"]` style access
- **`ManyToOneMultidimensionalIndexerMap<T>`**: For `array[0, 1]` style access

### 4. Reading Layer

#### Cell Readers

Readers extract raw values from Excel cells:

**Single Cell Readers** (`ICellReader`):
- **`ColumnIndexReader`**: Reads by zero-based column index
- **`ColumnNameReader`**: Reads by column name (requires heading)

**Multi-Cell Readers** (`ICellsReader`):
- **`ColumnIndicesReader`**: Reads multiple columns by index
- **`ColumnNamesReader`**: Reads multiple columns by name
- **`ColumnsMatchingReader`**: Reads columns matching a pattern (regex or predicate)
- **`AllColumnNamesReader`**: Reads all named columns

#### Reader Factories

Factories create readers lazily per sheet (for performance):

**Single Cell Factories** (`ICellReaderFactory`):
- `ColumnIndexReaderFactory`
- `ColumnNameReaderFactory`

**Multi-Cell Factories** (`ICellsReaderFactory`):
- `ColumnIndicesReaderFactory`
- `ColumnNamesReaderFactory`
- `ColumnsMatchingReaderFactory`
- `CharSplitReaderFactory` / `StringSplitReaderFactory`: Split a single cell value
- `AllColumnNamesValueReaderFactory`: Read all columns

### 5. Transformation Pipeline

The `ValuePipeline<T>` processes cell values through a multi-stage pipeline:

#### Stage 1: Value Reading

Raw value is read from Excel cell via reader:
- Returns `ReadCellResult` containing column index and string value
- Respects `PreserveFormatting` flag for formatted strings

#### Stage 2: Transformation (`ICellTransformer`)

String value is transformed before mapping:
- **`TrimCellValueTransformer`**: Trims whitespace
- Custom transformers can be added via `AddCellValueTransformer()`

#### Stage 3: Mapping (`ICellMapper`)

String value is converted to target type via mapper pipeline:

**Built-in Mappers**:
- **`StringMapper`**: Identity mapper for strings
- **`ChangeTypeMapper`**: Uses `IConvertible` for primitives (int, double, decimal, etc.)
- **`BoolMapper`**: Parses boolean values with various formats
- **`DateTimeMapper`**: Parses DateTime from Excel OLE Automation dates or strings
- **`GuidMapper`**: Parses GUID strings
- **`EnumMapper`**: Parses enum values by name (with optional case-insensitivity)
- **`UriMapper`**: Parses URI strings
- **`DictionaryMapper<T>`**: Maps nested objects using sub-maps
- **`ConvertUsingMapper`**: Custom conversion via delegates

**Mapper Pipeline Behavior**:

Each mapper returns a `CellMapperResult` with an action:
1. **`IgnoreResultAndContinueMapping`**: Skip this mapper, try next
2. **`UseResultAndContinueMapping`**: Store result, but try next mapper too
3. **`UseResultAndStopMapping`**: Use this result and stop

The pipeline processes mappers in order until one returns "stop" or the list ends.

#### Stage 4: Fallback Handling (`IFallbackItem`)

If mapping fails or value is empty, fallbacks provide default values:

**Empty Fallback** (`EmptyFallback`):
- Triggered when cell value is empty/null
- **`FixedValueFallback`**: Returns a constant value
- **`ThrowFallback`**: Throws `ExcelMappingException`

**Invalid Fallback** (`InvalidFallback`):
- Triggered when all mappers fail
- Same implementations as empty fallback

**Fallback Strategy** (`FallbackStrategy`):
- **`ThrowIfPrimitive`**: Throw exception for primitives, use default for reference types
- **`SetToDefaultValue`**: Use type's default value (0, null, false, etc.)

### 6. Auto-Mapping Layer

**Entry Point**: `AutoMapper` utility class

When no explicit map is registered, AutoMapper automatically creates maps using reflection:

#### Type Support

**Primitive Types**:
- Numeric: `int`, `long`, `double`, `decimal`, `byte`, `short`, `float`, etc.
- Text: `string`, `char`
- Other: `bool`, `DateTime`, `Guid`, `Uri`, `Enum`
- Nullable versions of all value types

**Collection Types**:
- Arrays: `T[]`
- Lists: `List<T>`, `IList<T>`, `ICollection<T>`, `IEnumerable<T>`
- Sets: `HashSet<T>`, `ISet<T>`, `FrozenSet<T>`, `ImmutableHashSet<T>`
- Dictionaries: `Dictionary<TKey, TValue>`, `IDictionary<TKey, TValue>`, `FrozenDictionary<TKey, TValue>`, `ImmutableDictionary<TKey, TValue>`

**Complex Types**:
- Classes with public properties/fields
- Record types
- Dynamic objects (via `ExpandoObject`)

#### Auto-Mapping Process

1. **Discover Members**: Find all public instance properties and fields
2. **Apply Attributes**: Process `ExcelColumn*` attributes for each member
3. **Create Maps**: Generate appropriate map type (OneToOne, ManyToOne, etc.)
4. **Setup Pipelines**: Configure mappers and fallbacks for each property type
5. **Register**: Cache the generated map in configuration

#### Expression Mapping

**Entry Point**: `ExpressionAutoMapper` utility class

Supports fluent mapping syntax via expression trees:

```csharp
var map = new ExcelClassMap<Person>();
map.Map(p => p.Name);                          // OneToOneMap
map.Map(p => p.Addresses);                     // ManyToOneEnumerableMap
map.Map(p => p.Addresses[0]);                  // ManyToOneEnumerableIndexerMap
map.Map(p => p.PhoneNumbers["home"]);          // ManyToOneDictionaryIndexerMap
map.Map(p => p.Matrix[0, 1]);                  // ManyToOneMultidimensionalIndexerMap
```

**Expression Parsing**:
- Walks expression tree to extract property/field/indexer access
- Handles chained member access: `p => p.Address.City`
- Handles indexer access: `p => p.List[0]`, `p => p.Dict["key"]`
- Skips `Convert` expressions for type compatibility
- Validates indexer types and constant indices

**Type Safety**:
- Validates indexer argument types (int for arrays/lists, constant keys for dictionaries)
- Rejects non-constant indexer arguments at parse time
- Ensures proper type mapping between Excel data and property types

### 7. Attribute System

Attributes control mapping behavior declaratively:

**Column Specification**:
- **`[ExcelColumnName("Name")]`**: Map to column by name
- **`[ExcelColumnIndex(0)]`**: Map to column by index
- **`[ExcelColumnNames("First", "Last")]`**: Map to multiple columns by name
- **`[ExcelColumnIndices(0, 1, 2)]`**: Map to multiple columns by index
- **`[ExcelColumnsMatching("Regex.*")]`**: Map to columns matching pattern
- **`[ExcelColumnMatching("Regex")]`**: Map to single column matching pattern

**Behavior Modifiers**:
- **`[ExcelOptional]`**: Don't throw if column is missing
- **`[ExcelIgnore]`**: Skip this property during auto-mapping
- **`[ExcelDefaultValue(value)]`**: Use default value if cell is empty
- **`[ExcelPreserveFormatting]`**: Read formatted string value instead of raw value

### 8. Factory System

Factories create collection instances during mapping:

#### Enumerable Factories (`IEnumerableFactory<T>`)

- **`ArrayEnumerableFactory<T>`**: Creates `T[]`
- **`ConstructorEnumerableFactory<T>`**: Uses `new()` constructor
- **`ConstructorSetEnumerableFactory<T>`**: Uses `new()` for sets
- **`HashSetEnumerableFactory<T>`**: Creates `HashSet<T>`
- **`FrozenSetEnumerableFactory<T>`**: Creates `FrozenSet<T>`
- **`ImmutableHashSetEnumerableFactory<T>`**: Creates `ImmutableHashSet<T>`
- **`AddEnumerableFactory<T>`**: Uses `.Add()` method
- **`ICollectionTImplementingEnumerableFactory<T>`**: For `ICollection<T>` implementations

#### Dictionary Factories (`IDictionaryFactory<TKey, TValue>`)

- **`DictionaryFactory<TKey, TValue>`**: Creates `Dictionary<TKey, TValue>`
- **`ConstructorDictionaryFactory<TKey, TValue>`**: Uses `new()` constructor
- **`FrozenDictionaryFactory<TKey, TValue>`**: Creates `FrozenDictionary<TKey, TValue>`
- **`ImmutableDictionaryFactory<TKey, TValue>`**: Creates `ImmutableDictionary<TKey, TValue>`
- **`AddDictionaryFactory<TKey, TValue>`**: Uses `.Add()` method

#### Multidimensional Array Factories (`IMultidimensionalArrayFactory`)

- **`TwoDimensionalArrayFactory<T>`**: Creates `T[,]` arrays
- **`MultidimensionalArrayFactory<T>`**: Creates arrays with arbitrary dimensions

## Data Flow Example

### Simple One-to-One Mapping

```
Excel File → ExcelImporter → ExcelSheet → ReadRows<Person>()
    ↓
ExcelClassMap<Person>.TryGetValue()
    ↓
OneToOneMap<string>.TryGetValue() for "Name" property
    ↓
ColumnNameReaderFactory.GetCellReader() → ColumnNameReader
    ↓
ColumnNameReader.TryGetValue() → ReadCellResult("John Doe")
    ↓
ValuePipeline<string>:
  1. TrimCellValueTransformer → "John Doe"
  2. StringMapper.Map() → CellMapperResult.Success("John Doe")
    ↓
Person.Name = "John Doe"
```

### Complex Many-to-One Mapping

```
Excel: Columns "Score1", "Score2", "Score3" → List<int> Scores
    ↓
ManyToOneEnumerableMap<int>.TryGetValue()
    ↓
ColumnNamesReaderFactory.GetCellsReader() → ColumnNamesReader
    ↓
ColumnNamesReader.TryGetValues() → [ReadCellResult("95"), ReadCellResult("87"), ReadCellResult("92")]
    ↓
For each ReadCellResult:
  ValuePipeline<int>:
    1. ChangeTypeMapper.Map() → CellMapperResult.Success(95)
    2. ListEnumerableFactory<int>.Add(95)
    ↓
ListEnumerableFactory<int>.End() → List<int> {95, 87, 92}
    ↓
Person.Scores = [95, 87, 92]
```

## Error Handling

### Exception Types

- **`ExcelMappingException`**: Base exception for mapping errors
  - Includes sheet name, row index, column information
  - Thrown when required columns are missing
  - Thrown when type conversion fails without fallback

### Error Recovery

1. **Optional Properties**: Use `[ExcelOptional]` or `.WithOptional()` to skip missing columns
2. **Fallback Values**: Use `[ExcelDefaultValue]` or `.WithEmptyFallback()` / `.WithInvalidFallback()`
3. **Nullable Types**: Automatically map to `null` on empty/invalid values

## Performance Optimizations

1. **Reader Caching**: `ICellReader` instances are cached per sheet (not per row)
2. **Map Registration**: Class maps are registered once and reused
3. **Lazy Evaluation**: `ReadRows<T>()` uses `yield return` for streaming
4. **Type Reflection Caching**: Property/field info is cached during map creation

## Extensibility Points

1. **Custom Mappers**: Implement `ICellMapper` for custom type conversions
2. **Custom Transformers**: Implement `ICellTransformer` for string preprocessing
3. **Custom Fallbacks**: Implement `IFallbackItem` for custom default value logic
4. **Custom Factories**: Implement `IEnumerableFactory<T>` or `IDictionaryFactory<TKey, TValue>`
5. **Custom Readers**: Implement `ICellReader` or `ICellsReader` for custom column reading logic
6. **Custom Column Matchers**: Implement `IExcelColumnMatcher` for advanced column selection

## Type System

### Supported Types

#### Value Types
- Primitives: `int`, `long`, `double`, `decimal`, `float`, `byte`, `short`, `uint`, `ulong`, `ushort`, `sbyte`
- Special: `bool`, `char`, `DateTime`, `Guid`
- Enums: Any enum type (with optional case-insensitive parsing)
- Nullable: All nullable value types

#### Reference Types
- `string`, `object`, `Uri`
- Collections: Any `IEnumerable<T>` implementation
- Dictionaries: Any `IDictionary<TKey, TValue>` implementation
- Complex: Classes, records, interfaces (via auto-mapping)

#### Generic Constraints
- Dictionary keys must be `notnull` (`TKey : notnull`)
- Factory types must have parameterless constructors or be well-known types