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

The public API is the importer, sheets and regions, class maps and the fluent methods on their maps, attributes, and four extension points (`ICellConverter<T>`, `ICellTransformer`, `ICellFallback<T>` and `IColumnSelector`). Everything else, including `IMap` (the interface every map implements), `ClassMapMember` (a class map's member and its map), `AutoMapper`, the readers, mappers, fallbacks, collection builders and bindings, is internal, so it can change without breaking users.

## Reading sheets

| Type | Role |
|---|---|
| `ExcelImporter` | Opens a workbook or CSV file and owns its `IExcelDataReader` and `ExcelImporterConfiguration`. `ReadSheet`/`TryReadSheet` return sheets by position or name. |
| `RowSource` (internal) | Owns the reader's position. The reader is forward-only and shared by every sheet of the importer, so this is the only code that calls `Read`, `NextResult` and `Reset`: sheets ask it for the sheet and row they need, and it reads forward or rewinds. Blank rows at the end of a sheet aren't rows: when the reader reaches a blank row it reads ahead to the next row with a value, and the blank rows before it are read through a `BlankRowReader` (whose cells are all empty) while the reader stays ahead; if only blank rows remain, the sheet ends. It also counts rows, by reading them, when a range counts from the end. |
| `ExcelSheet` | A worksheet and a position in it. Reads the heading (`ReadHeading`, `FindHeading`), then rows (`ReadRows<T>`, `ReadRow<T>`, `ReadRowResults<T>`). It keeps its own position, so two `ExcelSheet` objects for one worksheet don't interfere. Not thread-safe. It's the public front of a `SheetView` (internal), which does the reading: it holds the data range and heading and maps rows, with `Worksheet` (what's fixed about the worksheet, such as its name, index and reader), `RowCursor` (the row it's on, and moving to the next row of data), `SheetBindings` (the columns, readers and bound class maps worked out for it) and `HeadingSearch` (`FindHeading`). Internal code, such as maps, readers and `CellContext`, works with the view. |
| `ExcelHeading` | The heading row: column names and their indexes. `ColumnUtilities` finds a column by name: exactly (ignoring case by default), then ignoring white space around the names, then, when `IgnoreSeparatorsInColumnNames` is set, ignoring separators; two columns that match the same way are ambiguous. |
| `ExcelRegion` | An immutable block of a worksheet with its own heading choice (`WithHeadingAt`, `WithoutHeading`, `WithHeadingSearch`). Each read creates a `SheetView` of the region's `Worksheet`, and reads (or finds) its heading when the method is called, as a sheet's methods do, with its own `DataRange`, heading, cursor and bindings, so regions use the same reading, binding and heading code as sheets. The view's `Sheet` is the user's sheet, which converters, column selectors and errors get, with the region's heading (`CellContext.Heading`, `ColumnSelectorContext.Heading`). Columns chosen by position count from the region's first column (`SheetView.ColumnOffset` and `ToSheetColumn`). |

`ExcelImporterConfiguration` holds the registered class maps and the settings that apply to every sheet: `Culture`, `SkipBlankLines`, `ValidateDataAnnotations`, `IgnoreSeparatorsInColumnNames` and `MaxColumnsPerSheet`.
A sheet takes the settings when it starts reading (`ReadSettings`), so changing them part way through a sheet only affects sheets read afterwards.

## Class maps

An `ExcelClassMap<T>` lists the members of `T` to map (its internal `Members`, each a `ClassMapMember`: the property or field, the map that reads it, and whether `Map(...)` mapped it). Class maps are created in one of two ways:

- **Explicitly**, by deriving from `ExcelClassMap<T>` (or passing a callback to `RegisterClassMap<T>`) and calling `Map(...)`. `ExpressionAutoMapper` (internal) walks the expression: a member (`p => p.Name`), a chain of members (`p => p.Address.City`, which creates nested class maps), or an indexer whose argument is known when the class map is created: a constant, or a variable or field the lambda captures (`p => p.Values[0]`, `p => p.Values[i]`, `p => p.Scores["Maths"]`, `p => p.Grid[0, 1]`). Mapping a member `Map(...)` already mapped, to the same kind of map, returns its map, so settings add up; a member auto-mapping mapped gets a new one. A member whose type `Map(...)` can't read from a cell, such as a `HashSet<T>`, records why (`CellPipeline.NoConversion`), for the error when the class map is registered or first used.
- **Automatically**, the first time a type without a registered map is read. `AutoMapper` finds the public properties and fields, and `MemberMapper` applies the member's attributes (`[ExcelColumnName]`, `[ExcelConverter]`, `[ExcelOptional]`, ...), rejecting attributes that choose columns in different ways. Values attributes give, such as `[ExcelDefaultValue]`'s, are converted to the member's type there (`MemberAttributes.ConvertValue`). Explicit maps use the same code to set a member up before the fluent methods change it.

What a member's type supports is decided once per type, by its `TypeRule<T>` (`TypeRule.For(type)`): its built-in conversion, a `CollectionRule<TElement>` if it has elements and a `DictionaryRule<TKey, TValue>` if it has keys and values. Each rule creates and sets up the maps for its type: collection, dictionary and indexer maps, and the collection or dictionary itself. `AutoMapper.TryCreateMap<T>` tries them in order: a value (the type's conversion, or one the member's attributes give), a dictionary, a collection, then an object with its own members. The rules are the only place auto-mapping makes generic types at run time, so the rest of it calls them without reflection.

Each member's map is one of:

| Map | Used for | Reads |
|---|---|---|
| `ValueMap<T>` | Single values: numbers, strings, dates, enums, `object`, and any type with a converter | One cell |
| `CollectionMap<TElement>` | Arrays, lists, sets and other collections | Several columns, or one cell split by separators |
| `DictionaryMap<TKey, TValue>` | Dictionaries, `ExpandoObject` and extra columns (`MapExtraColumns`) | Several columns, keyed by column name |
| `CollectionIndexerMap<T>`, `DictionaryIndexerMap<TKey, TValue>`, `MultidimensionalIndexerMap<T>` (internal, all `IndexerMap<TKey>`) | Elements mapped one by one: `Map(p => p.Values[0])` | One map per element |
| `ExcelClassMap<T>` | Nested objects (`MapObject(p => p.Address)`, `Map(p => p.Address.City)`, or auto-mapping) | The nested type's own members |
| `RowNumberMap` | `[ExcelRowNumber]` and `MapRowNumber` | The row's number |

Class maps and the indexer maps are made of other maps (`IParentMap`: their children, each with a member, index or key). `MapTree` walks the tree, so code that needs every map that reads cells, such as `ValidateHeading`, `FindHeading`, extra columns and checking a class map when it's registered or first used, doesn't switch over the kinds of map itself.

A single-value map (`ValueMap<T>`) converts its cells with an internal `CellPipeline<T>` (see [The value pipeline](#the-value-pipeline)). The collection and dictionary maps hold an `ElementMap<T>` for their elements or values (`WithElementMap`, `WithValueMap`), which has its own pipeline. Both maps derive from the public `ValuePipeline<T, TSelf>`, which holds the pipeline (`Pipeline`) and whose value methods configure it and return `TSelf`, so a chain of value methods keeps the map's type: `Map(p => p.Name).WithTrim().MakeOptional()`. The pipeline classes themselves (`CellPipeline`, `CellPipeline<T>`, in `src`) are internal; their names differ from `ValuePipeline<T, TSelf>` so the two aren't confused.

### Choosing columns

Each map that reads cells has a `ColumnSettings` (internal, in `src/Readers`): its `Columns`, whether they can be missing (`Optional`), and whether cells are read as formatted text (`PreserveFormatting`). It also handles a missing column, skipping an optional member or throwing the error. A `Columns` is the columns themselves: by name, by index or letter, by several names or indexes, by a predicate or regex, by an `IColumnSelector`, all columns, the columns no other member reads (`ExtraColumns`), or one cell's text split by separators (`SplitColumns`, with a `CellSplitter`). The fluent methods (`WithColumnName`, `WithColumnIndices`, `WithSeparators`, ...) and attributes choose it.

The columns are found once per sheet (`SheetView.FindColumn`, cached by `SheetBindings`): a single value reads the first column found, by its index; a collection or dictionary reads all of them through an `ICellsReader` (`ColumnIndicesReader`, or `SplitCellsReader` for a split cell). Each kind of `Columns` also describes itself: the `ExcelMappingException` for missing columns (naming the column, listing the heading's columns, and suggesting one whose name differs only by separators; an optional member is skipped instead), its text in `ValidateHeading`, and whether it looks columns up by name, which `FindHeading` needs.

## Binding

The first time a class map is used with a sheet, `SheetView.GetBoundMembers` (through its `SheetBindings`) binds each member of an `ExcelClassMap<T>` to the sheet as a `BoundMember<T>`, which keeps what's been worked out for that sheet. Bound members set the row through a `ref T`, with `ExcelMemberSetter<TOwner, TValue>` setters (generated, or compiled by `CompiledAccessors`), so a struct row is created and set in place and never boxed. The kind of bound member depends on the map's internal interface: maps whose value has a known type when bound are `IBindableMap`s (`ValueMap<T>` and nested class maps), and maps that read their own `Columns` are `IColumnsMap`s: single values, or `ICellsMap`s (collections and dictionaries).

- `BoundValueMember<TOwner, T>` (single values): keeps the member's column, and the value stays a `T` from the pipeline to the member, without boxing.
- `BoundClassMember<TOwner, T>` (nested class maps): reads the nested `T` with its class map, so a nested struct isn't boxed either.
- `BoundCellsMember<TOwner>` (collections and dictionaries): keeps the cells reader. The collection is built by the map's `CollectionBuilder<T>` or `DictionaryBuilder<TKey, TValue>`, chosen once for the member's type by its `CollectionRule`/`DictionaryRule`. Builders hold no state (the collection being built is a local), so one builder serves every sheet and thread without a lock. Collections created from a copy of their elements, such as immutable and frozen ones, gather them in a list each thread reuses.
- `BoundMapMember<TOwner>` (everything else, such as indexers and row numbers): maps through `IMap.TryGetValue`.

Bindings belong to the sheet, so a class map shared between importers or threads holds no per-sheet state.

A class map is frozen (`IMap.Freeze`) the first time it's used with a sheet (`SheetView.UseClassMap`, once per read rather than for every row), after checking every member can convert its cells, with every map under it: its members, nested class maps, indexer maps, and the pipelines of values and elements. Changing a frozen map throws an `InvalidOperationException` (`ThrowHelpers.ThrowIfFrozen`, and `NonNullCollection<T>` for the lists of members, mappers and transformers). So what's worked out from a map never goes stale and nothing checks for changes: a pipeline's plan is made when it's frozen, a class map's construction plan and a sheet's bindings are made once, and bound members find their column or cells reader once.

### Creating objects

`ExcelClassMap<T>` reads rows with a `ClassMapReader<T>`, made when the first row is read, which decides once how to create the type (a `ConstructionPlan<T>`, in `src/Binding`):

- With a **parameterless constructor**, it creates the object, then sets each member.
- Otherwise (records, and classes with read-only properties), it finds the public **constructor** whose parameters match mapped members by name and type (or have default values), reads every member first, and passes them to the constructor. Members that aren't parameters are set afterwards.

`CompiledAccessors` compiles setters and constructor calls once per member or constructor, and falls back to reflection where code can't be compiled (which boxes). A setter or constructor an `ExcelMappingContext` registered (`RegisteredAccessors`, written by the source generator as ordinary code: `ExcelMemberSetter<TOwner, TValue>` and `ExcelConstructor<TOwner>`, whose `ExcelConstructorArguments<TOwner>` reads each argument typed) is used before either, so native AOT apps don't use reflection. A row read as a single value (`ReadRows<int>()`, or `Map(p => p)`) is read by its `ValueMap<T>` as a `T`.

## The value pipeline

A `CellPipeline<T>` turns a cell into a `T`. The stages, in order:

1. **Transformers** change the cell's value, for example to trim it (`WithTrim`, `WithTransformers`, `[ExcelTransformer]`).
2. **Empty cells** use the empty fallback: by default `null` for reference and nullable types, and an error for other value types (`WithEmptyFallback`, `[ExcelDefaultValue]`, `[ExcelEmptyFallback]`).
3. **Mappers** are tried in order. Each one converts the value, says it's invalid, or doesn't handle it, so the next one tries. By default this is the built-in conversion for the member's type; `WithConverter` replaces it, `WithMapping` adds a mapping dictionary, and `WithFormats` and the other format methods configure it.
4. **Invalid values** use the invalid fallback: by default an `ExcelMappingException` for the cell (`WithInvalidFallback`, `[ExcelInvalidValue]`, `[ExcelInvalidFallback]`).

Every stage is typed, so values are converted as `T` from the cell to the member, without boxing:

- **Mappers** implement the internal `ICellMapper<T>`: `Convert(cell, context, out T value, ref error)` returns `Success`, `SuccessUnlessReplaced` (strings, so that a later mapping dictionary can replace the text), `Invalid` or `NotHandled`, and `CreateException(cell, context, error)` (from `ICellExceptionSource`, the non-generic part of the interface) creates the exception for an invalid cell, only when a fallback, `ReadRowResults` or an `IExcelErrorSink` uses it. The mappers for value types also implement `ICellMapper<T?>`, so nullable members share them.
- **Options** for the built-in conversion (formats, format provider, number style, URI kind and offset: `WithFormats`, `[ExcelFormats]` and so on) are the pipeline's `ConversionOptions`, which reach the mappers with each cell (`CellContext.Options`). The built-in mappers hold no state, so `TypeRule<T>` gives every member of a type the same ones. Setting an option after `WithConverter` adds the built-in mapper back, after the converter.
- **Converters** (`ICellConverter<T>`, and `WithConverter(value => ...)` through `DelegateCellConverter<T>`) are mappers too, through `CellConverterMapper<T>`. A mapper can keep what went wrong in `error`, which the pipeline passes back to `CreateException`: a converter's error can't be asked for afterwards without converting again.
- **Fallbacks** are `ICellFallback<T>`: the public interface, `FixedValueFallback<T>` for fixed values (a `T`: attribute values are converted when the class map is created) and `ThrowFallback<T>` for errors, which throws the cell's exception (made by `CellErrors`). While errors are being collected, the pipeline doesn't call a `ThrowFallback<T>`: it adds the cell to the `ErrorTarget` instead.
- **Transformers** are `ICellTransformer`, the public interface, including the built-in `TrimStringCellTransformer`.

The pipeline creates one `CellContext` for each cell and gives it to every transformer, mapper and fallback. When a pipeline is frozen, it takes a snapshot of its stages as arrays (its plan), so mapping a cell doesn't enumerate lists through interfaces.

`NullableCellConverter<T>` and `NullableCellFallback<T>` let a converter or fallback for `T` serve a `T?` member. For one given by an attribute, `TypeRule<T>.NullableAdapter` (an `UnderlyingTypeAdapter<T>`, made once per nullable type) wraps it, so a member's attributes are applied without reflection. The objects attributes give the type of are created by `MemberAttributes.CreateInstance`, which reports a type that can't be created with an `ExcelMappingException` naming the attribute and member.

| Member type | Built-in mapper (in `src/Mappers`) |
|---|---|
| `string` | `StringMapper` |
| Numbers | `StructNumberMapper<T>`/`INumberBaseMapper<T>` (.NET 8), `NumberMapper<T>` (netstandard); a number cell out of its type's range is found invalid by a range check (`MapperUtilities.IsInRange`), not a caught exception |
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
- **Collecting errors.** `ReadRowResults<T>` and `ReadRows<T>(IExcelErrorSink)` pass an `ErrorTarget` (in `src/Reading`) down while they map a row, through `IMap.TryGetValue`, the bound members and `CellPipeline<T>.Map`: errors are added to it rather than thrown, so every invalid cell in the row is reported, along with data-annotation failures (`ValidateDataAnnotations`), of the row's object and the nested objects its class map reads. A member whose value had an error isn't set, so it keeps its initial value, and a row with errors isn't returned. `ReadRowResults` uses an `ExceptionListTarget`, which creates each cell's exception. A sink uses a `SinkTarget`, which gives it an `ExcelMappingError`, a `ref struct` holding the cell and the mapper that rejected it, so the exception (`CellErrors.Create`) is only made if the sink calls `ToException()`.

## Trimming

The net8.0 build is annotated for trimming: the mapped type `T` and the types given in attributes are kept, and the reflection-based APIs are marked `[RequiresUnreferencedCode]`/`[RequiresDynamicCode]`. Generic mapper and adapter types created at run time are kept with `[DynamicDependency]`. `tests/TrimmingTests` is published trimmed and run in CI. Native AOT only has the code of the generic types it sees when compiling, so a type's `TypeRule<T>` and the built-in number mappers are created directly where the type is known (as for class maps and their `Map(...)` members), and generic types for types found at run time are made by `GenericTypes.Create`, with an instantiation an `ExcelMappingContext` registered (through `ExcelTypeRegistrar`) or else `MakeGenericType`. So a context's class maps are auto-mapped as usual, with native AOT; the source generator in `gen` writes a context's registrations from `[ExcelMappable]` attributes: `tests/AotTests` is published with native AOT and run in CI. See [Trimming and Native AOT](trimming.md).

## Public API

The public API of each target framework is listed in `src/PublicAPI/<framework>/PublicAPI.Unshipped.txt` (and `PublicAPI.Shipped.txt` after a release), checked by Microsoft.CodeAnalysis.PublicApiAnalyzers on every build, so a change to the public API fails the build until the lists change with it. The frameworks' APIs differ: .NET 8 also has `DateOnly`, `TimeOnly` and the generic `INumberBase<T>` options. Package validation separately compares the package with the last release on NuGet (see `CLAUDE.md`).

## Source layout

| Path | Contents |
|---|---|
| `src/*.cs` | The public types: importer, sheets, regions, class maps and maps, attributes, extension points, `ValuePipeline<T, TSelf>`, `ElementMap<T>` and the type-specific value methods in `ValuePipelineExtensions` |
| `src/Maps` | The internal maps: `ClassMapMember` (a member of a class map and its map), the indexer maps (`IndexerMap<TKey>` and its kinds) and `RowNumberMap` |
| `src/Abstractions` | The internal map interfaces (`IMap`, `IParentMap`, `IColumnsMap`, `ICellsMap`, `IBindableMap`), `ICellMapper<T>`, and `ReadCellResult`, the cell a mapper converts |
| `src/Reading` | Reading sheets: `SheetView`, `Worksheet`, `RowSource` and `BlankRowReader` (blank rows the reader has passed), `RowCursor`, `HeadingSearch`, `ReadSettings`, and `ErrorTarget` (where collected errors go) |
| `src/AutoMapping` | Creating maps: `AutoMapper`, `ExpressionAutoMapper`, the rules (`TypeRule`, `CollectionRule`, `DictionaryRule`), `MemberMapper` and `MemberAttributes` (a member's attributes) and `UnderlyingTypeAdapter` |
| `src/Binding` | Binding class maps to sheets and reading rows with them: `ClassMapReader<T>`, `SheetBindings`, `BoundMember<T>` and its kinds, `ConstructionPlan<T>` (how a class map creates its type), `CompiledAccessors` (compiled setters and constructors) and `ConstructorUtilities` (matching constructors to members) |
| `src/Readers` | `ColumnSettings`, `Columns` and its kinds (a file each), splitting cells (`CellSplitter`), and the readers of several cells (`ICellsReader`, `ColumnIndicesReader`, `SplitCellsReader`) |
| `src/Mappers`, `src/Transformers`, `src/Fallbacks` | Built-in conversion, transformers and fallbacks, and the adapters for the public extension points |
| `src/Builders` | `CollectionBuilder<T>` and `DictionaryBuilder<TKey, TValue>`: how each collection and dictionary type is built |
| `src/Utilities` | Helpers used across the library: walking and checking map trees (`MapTree`, `MapValidation`), `ConversionOptions`, `NonNullCollection<T>`, argument checks (`ThrowHelpers`), cell error messages (`CellErrors`), C# type names for messages (`TypeNames`), and column, format, date, enum and reflection helpers |
| `tests/ExcelMapper` | Unit tests of public types, one file per type, mirroring `src/` |
| `tests/Maps`, `tests/Fallbacks`, `tests/Transformers` | Mapping tests, which read spreadsheets and cover the internal types through the public API |
| `tests/TrimmingTests` | A trimmed app that checks mapping works after trimming |
| `gen` | The source generator: `ExcelMappableGenerator` writes each `[ExcelMappable]` context's `RegisterTypes`, properties and `Default`, from the types `TypeWalker` finds by following auto-mapping's rules (a built-in conversion, a dictionary, a collection, or an object's public properties and fields), and warns (EXM005 to EXM007) about types those rules can't map |
| `tests/SourceGeneration` | Tests of the generator: what it writes and reports, and generated contexts mapping as auto-mapping does |
| `tests/AotTests` | A native AOT app that checks mapping contexts, and class maps of built-in types, work with native AOT without warnings |
| `benchmarks` | BenchmarkDotNet benchmarks (see [Performance](performance.md)) |
