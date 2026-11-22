# ExcelMapper Architecture

This page is for contributors: how a row of a spreadsheet becomes an object, and where each part of that lives in `src/`. For using the library, see the [README](../README.md) and the other pages in `docs/`.

## Overview

ExcelMapper reads spreadsheets with [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) and maps each row to an object. The work is split into four parts, each of which does as much as it can once, so that mapping a row does as little as possible:

```
            once per type             once per sheet              once per row
┌──────────────────────────┐ ┌──────────────────────────┐ ┌──────────────────────────┐
│ Class map                │ │ Binding                  │ │ Mapping                  │
│ ExcelClassMap<T>: which  │→│ BoundMember per member:  │→│ read each member's cells,│
│ members, which columns,  │ │ columns resolved, cell   │ │ convert them through the │
│ how values are converted │ │ readers and setters made │ │ member's pipeline, create│
│ (explicit or auto-mapped)│ │                          │ │ or fill the object       │
└──────────────────────────┘ └──────────────────────────┘ └──────────────────────────┘
                                         ↑
                           RowSource: the reader's position
```

The public API is the importer, sheets and regions, class maps and the fluent methods on their maps, attributes, and four extension points (`ICellConverter<T>`, `ICellTransformer`, `ICellFallback<T>` and `IColumnSelector`). Everything else, including `IMap` (the interface every map implements), `ExcelPropertyMap` (a class map's member and its map), `AutoMapper`, the readers, mappers, fallbacks, collection factories and bindings, is internal, so it can change without breaking users.

## Reading sheets

| Type | Role |
|---|---|
| `ExcelImporter` | Opens a workbook or CSV file and owns its `IExcelDataReader` and `ExcelImporterConfiguration`. `ReadSheet`/`TryReadSheet` return sheets by position or name. |
| `RowSource` (internal) | Owns the reader's position. The reader is forward-only and shared by every sheet of the importer, so this is the only code that calls `Read`, `NextResult` and `Reset`: sheets ask it for the sheet and row they need, and it reads forward or rewinds. It also counts rows, leaving out the empty lines at the end of CSV files. |
| `ExcelSheet` | A worksheet and a position in it. Reads the heading (`ReadHeading`, `FindHeading`), then rows (`ReadRows<T>`, `ReadRow<T>`, `ReadRowResults<T>`). It keeps its own position, so two `ExcelSheet` objects for one worksheet don't interfere. Not thread-safe. |
| `ExcelHeading` | The heading row: column names and their indexes, for finding columns by name (`ColumnUtilities` also matches names without separators when `IgnoreSeparatorsInColumnNames` is set). |
| `ExcelRegion` | An immutable block of a worksheet with its own heading choice (`WithHeadingAt`, `WithoutHeading`, `WithHeadingSearch`). Each read creates a private `ExcelSheet` for the region, with its own `DataRange`, heading and position, so regions reuse the sheet's reading, binding and heading code. Columns chosen by position count from the region's first column (`ExcelSheet.ColumnOffset` and `ToSheetColumn`). |

`ExcelImporterConfiguration` holds the registered class maps and the settings that apply to every sheet: `Culture` (set as the current culture while a row is mapped, so parsing uses it), `SkipBlankLines`, `ValidateDataAnnotations` and `IgnoreSeparatorsInColumnNames`.

## Class maps

An `ExcelClassMap<T>` lists the members of `T` to map (its internal `Properties`, each an `ExcelPropertyMap` of a member and its map). Class maps are created in one of two ways:

- **Explicitly**, by deriving from `ExcelClassMap<T>` (or passing a callback to `RegisterClassMap<T>`) and calling `Map(...)`. `ExpressionAutoMapper` (internal) walks the expression: a member (`p => p.Name`), a chain of members (`p => p.Address.City`, which creates nested class maps), or an indexer with a constant argument (`p => p.Values[0]`, `p => p.Scores["Maths"]`, `p => p.Grid[0, 1]`).
- **Automatically**, the first time a type without a registered map is read. `AutoMapper` finds the public properties and fields, and `MemberMapper` applies the member's attributes (`[ExcelColumnName]`, `[ExcelConverter]`, `[ExcelOptional]`, ...). Explicit maps use the same code to set a member up before the fluent methods change it.

Each member's map is one of:

| Map | Used for | Reads |
|---|---|---|
| `ValueMap<T>` | Single values: numbers, strings, dates, enums, `object`, and any type with a converter | One cell |
| `CollectionMap<TElement>` | Arrays, lists, sets and other collections | Several columns, or one cell split by separators |
| `DictionaryMap<TKey, TValue>` | Dictionaries, `ExpandoObject` and extra columns (`MapExtraColumns`) | Several columns, keyed by column name |
| `CollectionIndexerMap<T>`, `DictionaryIndexerMap<TKey, TValue>`, `MultidimensionalIndexerMap<T>` (internal) | Elements mapped one by one: `Map(p => p.Values[0])` | One map per element |
| `ExcelClassMap<T>` | Nested objects (`Map(p => p.Address)` or auto-mapping) | The nested type's own members |
| `RowNumberMap` | `[ExcelRowNumber]` and `MapRowNumber` | The row's number |

A single-value map is itself a `ValuePipeline<T>` (see [The value pipeline](#the-value-pipeline)). The collection and dictionary maps hold an `ElementMap<T>`, another pipeline, for their elements or values (`WithElementMap`, `WithValueMap`). Both derive from `ValuePipeline<T, TSelf>`, whose value methods return `TSelf`, so a chain of value methods keeps the map's type: `Map(p => p.Name).WithTrim().MakeOptional()`.

### Choosing columns

Each map has a reader factory (internal) that finds its columns in a sheet: by name, by index or letter, by several names or indexes, by a predicate or regex, all columns, the columns no other member reads, or one column's text split by separators. The fluent methods (`WithColumnName`, `WithColumnIndices`, `WithSeparators`, ...) and attributes choose the factory. `IColumnSelector` is the public way to choose columns in code: `ColumnSelectorReaderFactory` asks it once per sheet.

A factory creates a cell reader for a sheet once; the reader then reads the member's cells from each row. Missing columns throw an `ExcelMappingException` that names the column and suggests a similarly named one, unless the member is optional.

## Binding

The first time a class map is used with a sheet, `ExcelSheet.GetBoundMembers` binds each member to the sheet:

- `BoundOneToOneMember<T>` (single values): keeps the member's cell reader and a compiled `Action<object, T>` setter, so values stay typed from the pipeline to the member, without boxing.
- `BoundCellsMember` (collections and dictionaries): keeps the cells reader, and its own copy of the collection factory (`IBindingFactory`), so one sheet's reads don't need a lock.
- `BoundMapMember` (everything else, such as nested class maps, indexers and row numbers): maps through `IMap.TryGetValue`.

Bindings belong to the sheet, so a class map shared between importers or threads holds no per-sheet state. They're rebuilt if the class map's members or a map's columns change (both have version numbers).

### Creating objects

`ExcelClassMap` decides once how to create its type (a construction plan):

- With a **parameterless constructor**, it creates the object, then sets each member.
- Otherwise (records, and classes with read-only properties), it finds the public **constructor** whose parameters match mapped members by name and type (or have default values), reads every member first, and passes them to the constructor. Members that aren't parameters are set afterwards.

`CompiledAccessors` compiles setters and constructor calls once per member or constructor, and falls back to reflection where code can't be compiled.

## The value pipeline

A `ValuePipeline<T>` turns a cell into a `T`. The stages, in order:

1. **Transformers** change the cell's value, for example to trim it (`WithTrim`, `WithTransformers`, `[ExcelTransformer]`).
2. **Empty cells** use the empty fallback: by default `null` for reference and nullable types, and an error for other value types (`WithEmptyFallback`, `[ExcelDefaultValue]`, `[ExcelEmptyFallback]`).
3. **Converters** are tried in order. Each one converts the value, says it's invalid, or doesn't handle it, so the next one tries. By default this is the built-in conversion for the member's type; `WithConverter` replaces it, `WithMapping` adds a mapping dictionary, and `WithFormats` and the other format methods configure it.
4. **Invalid values** use the invalid fallback: by default an `ExcelMappingException` for the cell (`WithInvalidFallback`, `[ExcelInvalidValue]`, `[ExcelInvalidFallback]`).

Every stage is typed, so values are converted as `T` from the cell to the member, without boxing:

- **Mappers** implement the internal `ICellMapper<T>`: `Convert(cell, out T value)` returns `Success`, `SuccessUnlessReplaced` (strings, so that a later mapping dictionary can replace the text), `Invalid` or `NotHandled`, and `CreateException(cell)` creates the exception for an invalid cell, only when a fallback or `ReadRowResults` uses it. The mappers for value types also implement `ICellMapper<T?>`, so nullable members share them.
- **Converters** (`ICellConverter<T>`, and `WithConverter(value => ...)` through `DelegateCellConverter<T>`) are carried in the list of mappers by `CellConverterMapper<T>`, and called by the pipeline with the cell's `CellContext`. The pipeline keeps the error a converter returns, since it can't be asked for an exception afterwards without converting again.
- **Fallbacks** are `ICellFallback<T>`: the public interface, `FixedValueFallback<T>` for fixed values and `ThrowFallback<T>` for errors (which records the error instead of throwing while `ReadRowResults` collects errors).
- **Transformers** implement the internal `IPipelineTransformer`, which works on the cell directly: `TrimStringCellTransformer`, and `CellTransformerAdapter` for `ICellTransformer`.

`NullableCellConverter<T>` and `NullableCellFallback<T>` let a converter or fallback for `T` serve a `T?` member. The first time a pipeline maps a value, it takes a snapshot of its stages as arrays (taken again if the pipeline changes), so mapping a cell doesn't enumerate lists through interfaces.

| Member type | Built-in mapper (in `src/Mappers`) |
|---|---|
| `string` | `StringMapper` |
| Numbers | `StructNumberMapper<T>`/`INumberBaseMapper<T>` (.NET 8), `NumberMapper<T>` (netstandard) |
| `bool` | `BoolMapper` (also `1`/`0` and `yes`/`no`) |
| `DateTime`, `DateTimeOffset`, `DateOnly`, `TimeOnly`, `TimeSpan` | `DateTimeMapper`, `DateTimeOffsetMapper`, `DateOnlyMapper`, `TimeOnlyMapper`, `TimeSpanMapper` (Excel dates, serial numbers and text in the default or given formats) |
| Enums | `EnumMapper<T>` |
| `Guid`, `Uri`, `Version` | `GuidMapper`, `UriMapper`, `VersionMapper` |
| `object`, `IConvertible` members | `ObjectMapper<T>` (the value as Excel stores it) |
| Other `IParsable<T>` (.NET 8) and `IConvertible` types | `ParsableMapper<T>` and `StructParsableMapper<T>`, `ChangeTypeMapper<T>` |
| Values given by `[ExcelMappingDictionary]` and `WithMapping` | `MappingDictionaryMapper<T>` |

## Errors

- **`ExcelMappingException`** reports the cell (`CellAddress`, `ColumnName`, `Member`, one-based row in the message) and wraps the converter's exception.
- **Exceptions are created only when they're used.** A mapper reports an invalid cell without an exception, and the pipeline asks it for one (`ICellMapper<T>.CreateException`) only if the fallback uses it. A converter can return `CellConversionResult<T>.Invalid(() => exception)`. Fixed-value fallbacks don't use the exception, so an invalid cell with a fallback doesn't create one.
- **`ReadRowResults<T>`** sets `ExcelSheet.CollectedErrors` while it maps a row: errors are added to it rather than thrown, so every invalid cell in the row is reported, along with data-annotation failures (`ValidateDataAnnotations`). A member whose value had an error isn't set, so it keeps its initial value.

## Trimming

The net8.0 build is annotated for trimming: the mapped type `T` and the types given in attributes are kept, and the reflection-based APIs are marked `[RequiresUnreferencedCode]`/`[RequiresDynamicCode]`. Generic mapper and adapter types created at run time are kept with `[DynamicDependency]`. `tests/TrimmingTests` is published trimmed and run in CI. See [Trimming and Native AOT](trimming.md).

## Source layout

| Path | Contents |
|---|---|
| `src/*.cs` | The public types (importer, sheets, regions, class maps and maps, attributes, extension points, `ValuePipeline<T>`, `ValuePipeline<T, TSelf>`, `ElementMap<T>` and the type-specific value methods in `ValuePipelineExtensions`), and the internal `IMap`, `ExcelPropertyMap` and indexer maps |
| `src/Readers` | Reader factories and readers (internal) |
| `src/Mappers`, `src/Transformers`, `src/Fallbacks` | Built-in conversion, transformers and fallbacks, and the adapters for the public extension points (internal) |
| `src/Factories` | Collection and dictionary factories (internal) |
| `src/Abstractions` | The interfaces of the internal types, and the indexer map interfaces |
| `src/Utilities` | Auto-mapping (`AutoMapper`, `MemberMapper`, `ExpressionAutoMapper`), binding (`BoundMember`), `RowSource`, `CompiledAccessors`, constructor matching and parsing helpers |
| `tests/ExcelMapper` | Unit tests of public types, one file per type, mirroring `src/` |
| `tests/Maps`, `tests/Fallbacks`, `tests/Transformers` | Mapping tests, which read spreadsheets and cover the internal types through the public API |
| `tests/TrimmingTests` | A trimmed app that checks mapping works after trimming |
| `benchmarks` | BenchmarkDotNet benchmarks (see [Performance](performance.md)) |
