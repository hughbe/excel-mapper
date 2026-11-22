# Migrating to 4.0

ExcelMapper 4.0 fixes several long-standing bugs and makes mapping more predictable. Most code keeps working unchanged: reading rows, attributes, class maps and the fluent API are source compatible, but some **behaviour** changed, and the low-level extension types are replaced by simpler ones (see [Extension types](#extension-types)). This page lists what to check, most commonly affected first, and ends with [what's new](#new-in-40). See the [changelog](../CHANGELOG.md) for everything that changed.

## Checklist

1. [`object` members now receive typed values](#object-members-receive-typed-values)
2. [Text is read with the invariant culture](#text-is-read-with-the-invariant-culture)
3. [Fractional numbers can't be mapped to integers](#fractional-numbers-cant-be-mapped-to-integers)
4. [`WithConverter` replaces the built-in conversion](#withconverter-replaces-the-built-in-conversion)
5. [`ReadRow` reads the heading first](#readrow-reads-the-heading-first)
6. [Error messages use cell addresses](#error-messages-use-cell-addresses)
7. [Sheet names are case-insensitive](#sheet-names-are-case-insensitive)
8. [Extension types](#extension-types)
9. [Renamed types](#renamed-types)
10. [Renamed and removed methods](#renamed-and-removed-methods)
11. [Smaller changes](#smaller-changes)
12. [You can remove the code page registration](#you-can-remove-the-code-page-registration)

## `object` members receive typed values

**Who's affected:** code that maps to `object` or `IConvertible` members, `Dictionary<string, object>`, `ExpandoObject`/`dynamic`, or untyped collections such as `ArrayList`, and then treats the values as strings.

**Before:** every value was the cell's text: a number `2.5` arrived as `"2.5"` (or `"2,5"`, depending on the machine's culture), a date as culture-formatted text, a boolean as `"True"`.

**Now:** values arrive as Excel stores them: `double` for numbers, `DateTime` for dates, `bool` for booleans, `TimeSpan` for durations and `string` for text.

**What to do:**

```csharp
// 3.x: worked because every value was a string
var amount = (string)row["Amount"];

// 4.0: use the typed value...
var amount = (double)row["Amount"];

// ...or convert explicitly, choosing the culture
var text = Convert.ToString(row["Amount"], CultureInfo.InvariantCulture);
```

To keep receiving text for a member, map it to `string`, or preserve formatting (`[ExcelPreserveFormatting]` or `.MakePreserveFormatting()`) to get the text as displayed in Excel. CSV files are unaffected: every CSV cell is text.

## Text is read with the invariant culture

**Who's affected:** apps that read text cells (every cell of a CSV file, and text cells in a workbook) with numbers or dates in a culture other than the invariant one, such as `1,5` or `15.03.2024`, and rely on the machine's culture to parse them; and apps that map number or date cells to `string` members and expect the machine's format.

**Before:** text was parsed, and numbers and dates were converted to text, with the current culture of the thread reading the rows. The same file could give different values on different machines: on a machine set to German, `1.5` was read as fifteen.

**Now:** the importer's `Culture` defaults to `CultureInfo.InvariantCulture`, so results are the same on every machine. Cells Excel stores as numbers and dates aren't parsed, so they're unaffected unless mapped to text.

**What to do:** set the culture the file was written in:

```csharp
importer.Configuration.Culture = CultureInfo.GetCultureInfo("de-DE");

// Or, to keep 3.x's behaviour, the machine's culture:
importer.Configuration.Culture = CultureInfo.CurrentCulture;
```

A member can have its own culture with `.WithFormatProvider(...)` (see [Culture](types.md#culture)).

## Fractional numbers can't be mapped to integers

**Who's affected:** code that maps numeric cells holding fractions (such as `2.7`) to `int`, `long`, `byte` and other integer types.

**Before:** the fraction was silently dropped: truncated to `2` on .NET 8 and rounded to `3` on .NET Framework and other netstandard runtimes.

**Now:** the cell is invalid, the same as the text `"2.7"`. Whole numbers such as `2.0` still map.

**What to do:** if fractions are expected, map to `decimal` or `double`, or decide what an invalid value becomes:

```csharp
[ExcelInvalidValue(0)]
public int Quantity { get; set; }

// Or round explicitly
importer.Configuration.RegisterClassMap<Product>(c => c.Map(p => p.Quantity)
    .WithConverter(value => (int)Math.Round(double.Parse(value!, CultureInfo.InvariantCulture))));
```

To find affected rows without stopping at the first one, use `sheet.ReadRowResults<T>()`.

## `WithConverter` replaces the built-in conversion

**Who's affected:** code that calls `WithConverter` on a member whose type ExcelMapper converts itself (numbers, dates, enums, `bool`, `Guid` and so on), expecting the converter to only handle values the built-in conversion can't.

**Before:** the converter was added after the built-in conversion, so whether it ran depended on the member's type. For a number, the converter was skipped whenever the cell was a valid number. For a string, it was used for every cell:

```csharp
importer.Configuration.RegisterClassMap<Product>(c =>
{
    c.Map(p => p.Quantity).WithConverter(value => int.Parse(value!) * 2);  // "3" gave 3, not 6
    c.Map(p => p.Name).WithConverter(value => value!.ToUpperInvariant());  // "a" gave "A"
});
```

**Now:** the converter replaces the built-in conversion, and any earlier converter, so it's used for every cell that isn't empty, whatever the type: `"3"` gives `6`. Empty cells still use the empty fallback, and a converter that throws uses the invalid fallback.

**What to do:** if the converter was meant as a fallback for values the built-in conversion can't handle, handle both in the converter:

```csharp
importer.Configuration.RegisterClassMap<Product>(c => c.Map(p => p.Quantity).WithConverter(value =>
    int.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var quantity)
        ? quantity
        : ParseWords(value)));
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
try
{
    var employees = sheet.ReadRows<Employee>().ToList();
}
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

**Who's affected:** code that implements or uses ExcelMapper's low-level types directly: custom mappers, transformers, fallbacks, reader factories or column matchers, or code that reads a map's `Mappers`, `Properties` or `ReaderFactory`. Code that only uses attributes, class maps and the fluent methods compiles unchanged.

**Now:** there are four extension points, and the types they replace are internal:

| To... | 3.x | 4.0 |
|---|---|---|
| Convert a cell | `ICellMapper` with `WithMappers(...)` or `[ExcelMapper]`; `CellMapperResult`; `ReadCellResult`; `WithConverter(readResult => ...)` | [`ICellConverter<T>`](mapping.md#converters) with `WithConverter(converter)` or `[ExcelConverter(typeof(...))]`, or `WithConverter(value => ...)` with the cell's text |
| Change a cell before it's converted | `Abstractions.ICellTransformer`, `TrimStringCellTransformer` | [`ICellTransformer`](mapping.md#transformers) (in the `ExcelMapper` namespace) with `WithTransformers(...)` or `[ExcelTransformer]`, and `WithTrim()` or `[ExcelTrimString]` |
| Give empty or invalid cells a value | `IFallbackItem`, `FixedValueFallback`, `ThrowFallback`, `WithFallbackItem`, `WithEmptyFallbackItem`, `WithInvalidFallbackItem` | [`ICellFallback<T>`](mapping.md#computed-fallbacks) with `WithEmptyFallback(fallback)`/`WithInvalidFallback(fallback)` or `[ExcelEmptyFallback]`/`[ExcelInvalidFallback]`; or `WithValueFallback(value)`, `WithEmptyFallback(value)` and `WithThrowingFallback()` |
| Choose columns in code | `ICellReaderFactory`, `ICellsReaderFactory`, `WithReaderFactory(...)`, `IExcelColumnMatcher` | [`IColumnSelector`](mapping.md#column-selectors) with `WithColumns(selector)`, `[ExcelColumnMatching(typeof(...))]` or `[ExcelColumnsMatching(typeof(...))]` |
| Configure a member | `IOneToOneMap`, `IManyToOneMap`, `IToOneMap` and their extension classes; a pipeline's `Mappers`, `Transformers`, `EmptyFallback` and `InvalidFallback`; `OneToOneMap<T>.Pipeline` | The map `Map(...)` returns, and its fluent methods. The column methods (`WithColumnName`, `MakeOptional`, ...) are methods of the maps; the value methods are methods of the maps too, and the ones for particular types, such as `WithFormats`, are in `ValuePipelineExtensions` (was `IValuePipelineExtensions`) |
| Map members by hand | `IMap`, `ExcelPropertyMap`, `ExcelClassMap.Properties`, the maps' constructors, the indexer maps, `AutoMapper.TryCreateClassMap`, `RegisterClassMap(Type, IMap)` | `Map(...)` in a class map, including `Map(p => p.Values[0])` for elements, or attributes on the type's members. To start from the auto-mapped class map (`AutoMapper.TryCreateClassMap`), call `AutoMap()` on a class map, for example `RegisterClassMap<T>(c => { c.AutoMap(); c.Map(...); })`. `TryGetClassMap` returns an `ExcelClassMap` |
| Create collections | `IEnumerableFactory<T>`, `IDictionaryFactory<TKey, TValue>`, `IMultidimensionalArrayFactory<T>` | ExcelMapper creates the collection for the member's type |

`ExcelMapper.Mappers.MappingDictionaryMapperBehavior` is now `ExcelMapper.MappingDictionaryBehavior` (see [Renamed types](#renamed-types)).

**What to do:**

A custom mapper becomes a converter. A converter gets the cell's value as Excel stores it (`value.Kind`, `value.Text`, `value.Number`, ...) and where the cell is (`context.CellAddress`, `context.Member`, ...). It isn't called for empty cells, which use the empty fallback:

```csharp
// 3.x
public class PhoneMapper : ICellMapper
{
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var digits = new string(readResult.StringValue.Where(char.IsDigit).ToArray());
        return digits.Length == 10
            ? CellMapperResult.Success(digits)
            : CellMapperResult.Invalid(new FormatException("Phone numbers have 10 digits."));
    }
}

Map(c => c.Phone).WithMappers(new PhoneMapper());

// 4.0
public class PhoneConverter : ICellConverter<string>
{
    public CellConversionResult<string> Convert(CellValue value, CellContext context)
    {
        var digits = new string(value.Text!.Where(char.IsDigit).ToArray());
        return digits.Length == 10
            ? CellConversionResult<string>.Success(digits)
            : CellConversionResult<string>.Invalid(() => new FormatException("Phone numbers have 10 digits."));
    }
}

importer.Configuration.RegisterClassMap<Contact>(c => c.Map(x => x.Phone).WithConverter(new PhoneConverter()));  // or [ExcelConverter(typeof(PhoneConverter))]
```

`CellMapperResult.Ignore()` becomes `CellConversionResult<T>.NotHandled`, which lets the next converter try; if none converts the value, it's invalid. A converter replaces the built-in conversion (see [`WithConverter` replaces the built-in conversion](#withconverter-replaces-the-built-in-conversion)).

A custom fallback becomes a typed fallback:

```csharp
// 3.x
public class UnknownFallback : IFallbackItem
{
    public object PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception exception, MemberInfo member)
        => -1;
}

// 4.0
public class UnknownFallback : ICellFallback<int>
{
    public int Fallback(CellValue value, CellContext context, Exception? exception) => -1;
}

importer.Configuration.RegisterClassMap<Employee>(c => c.Map(e => e.Age).WithInvalidFallback(new UnknownFallback()));  // or [ExcelInvalidFallback(typeof(UnknownFallback))]
```

A reader factory or column matcher becomes a column selector, which chooses columns once for each sheet and describes them for error messages:

```csharp
// 4.0
public class QuarterColumns : IColumnSelector
{
    // Positions in the worksheet, as GetColumnIndex gives them.
    public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context)
        => context.Heading!.ColumnNames
            .Where(name => name.StartsWith("Q"))
            .Select(context.Heading.GetColumnIndex)
            .ToArray();

    public string Describe(ColumnSelectorContext context) => "the columns starting with Q";
}

importer.Configuration.RegisterClassMap<Report>(c => c.Map(r => r.Quarters).WithColumns(new QuarterColumns()));
```

If you created a map with its constructor, or implemented `IMap`, map the member with `Map(...)` in a class map instead.

## Renamed types

**Who's affected:** code that names these types, for example in a helper method's signature or a `using static`. Code that uses `var` and lambdas compiles unchanged.

| 3.x | 4.0 |
|---|---|
| `OneToOneMap<T>` | `ValueMap<T>` |
| `ManyToOneEnumerableMap<T>` | `CollectionMap<T>` |
| `ManyToOneDictionaryMap<TKey, TValue>` | `DictionaryMap<TKey, TValue>` |
| `FallbackStrategy.ThrowIfPrimitive`, `FallbackStrategy.SetToDefaultValue` | `EmptyValueStrategy.Error`, `EmptyValueStrategy.DefaultValue` |
| `[ExcelUri(UriKind.Relative)]` (`ExcelUriAttribute`) | `[ExcelUriKind(UriKind.Relative)]` (`ExcelUriKindAttribute`) |
| `ExcelImporterUtils.RegisterClassMapsInNamespace(importer, ...)` | `importer.Configuration.RegisterClassMapsInNamespace(...)` |
| `ExcelMapper.Mappers.MappingDictionaryMapperBehavior` | `ExcelMapper.MappingDictionaryBehavior` |
| `ConvertUsingSimpleMapperDelegate<T>` | `Func<string?, T>` |

**What to do:** replace the old names with the new ones.

## Renamed and removed methods

**Who's affected:** code that calls these methods or sets these properties. Most are one-line changes; the compiler points to each.

| 3.x | 4.0 |
|---|---|
| `Map(p => p.Status, ignoreCase: true)` | `Map(p => p.Status).WithIgnoreCase()`, which only compiles for enums (the old overload compiled for any struct and threw when used) |
| `.WithColumnNameMatching("Name", "Full Name")` (the first of several columns) | `.WithColumnNames("Name", "Full Name")`, as `[ExcelColumnNames]` means. `WithColumnNameMatching` still takes a predicate or regex. |
| `.WithValueFallback(value)` | `.WithFallback(value)` |
| `map.Optional = true`, `map.PreserveFormatting = true` | `map.MakeOptional()`, `map.MakePreserveFormatting()`; the properties are read-only |
| `collectionMap.Pipeline`, `dictionaryMap.Pipeline` | `.WithElementMap(e => ...)`, `.WithValueMap(v => ...)` |
| `configuration.RegisterClassMap(typeof(T), classMap)` | `configuration.RegisterClassMap(classMap)`, which registers the class map for its own type: the removed overload let a map be registered for a different type |
| `importer.RegisterClassMapsInNamespace(...)` | `importer.Configuration.RegisterClassMapsInNamespace(...)`, next to `RegisterClassMap` |
| `importer.Reader` | Removed: the importer's sheets share the reader, and moving it broke their positions. Pass your own `IExcelDataReader` to `new ExcelImporter(reader)` to configure it. |
| `new ExcelClassMap(typeof(T))`, or a class deriving from `ExcelClassMap` | `new ExcelClassMap<T>()`, or derive from `ExcelClassMap<T>`: the non-generic class map, now abstract with members only `ExcelClassMap<T>` implements, couldn't map any members |
| `new ExcelRange("A1:B10")` | `ExcelRange.Parse("A1:B10")` |
| Setting `Index`, `Indices`, `Name` or `Names` on a column attribute | Pass them to the attribute's constructor. The attributes are sealed, except `ExcelColumnIndexAttribute`. |
| `WithColumnNames` and `WithColumnIndices` taking an `IReadOnlyList<T>` | `params T[]` or `IEnumerable<T>` |

**What to do:** make the change in the table for each compile error.

## Smaller changes

- **`ExcelRange`**: `default(ExcelRange)`, and the value `TryParse` returns on failure, now mean every row and column (previously an empty range). `Parse` rejects addresses Excel doesn't accept: non-ASCII letters, signs (`1:+5`), columns past `XFD` and rows past 1,048,576. Use the `Range` or index constructors for larger CSV files.
- **Column offsets**: with a `DataRange` that doesn't start at column A, `ExcelHeading.GetColumnIndex` and `GetColumnName` use the column's position in the sheet (as cell readers always did). Before 4.0, values were read from the wrong columns in this case.
- **Split values**: `"a,b,"` split into an array or list now includes the trailing empty value, which maps through the empty-value fallback (`null` for strings; an error for non-nullable numbers unless you add a fallback). An empty cell gives an empty collection.
- **Registering after auto-mapping**: `RegisterClassMap<T>` now replaces a map that auto-mapping created for `T` (for example by an earlier `ReadRows<T>()`) instead of throwing. Registering twice yourself still throws.
- **Changing class maps after reading**: a class map, and the maps of its members, can't be changed once rows have been read with it, including by `ValidateHeading` or `FindHeading`. `Map(...)` or a fluent method called afterwards throws an `InvalidOperationException`. Before 4.0, the change applied to the following rows. Configure the class map before reading, or register a new one: `RegisterClassMap<T>` replaces an auto-mapped map, and a new importer can register a different map for the same type.
- **Enumerating rows more than once**: each enumeration of the result of `ReadRows<T>()` or `ReadRowResults<T>()` reads the same rows, from the row after the sheet's position when the method was called. Before 4.0, the result could only be enumerated once: a second enumeration continued where the first stopped, so `rows.Any()` followed by `rows.ToList()` lost the first row. To continue reading where you stopped, use `ReadRow<T>()`, or call `ReadRows<T>()` again.
- **Columns with the same name**: a member whose name more than one column in the heading matches (such as two columns named "Name", or "Name" and "name") throws an `ExcelMappingException` naming both columns, instead of silently reading the first. Choose one with `[ExcelColumnLetter]`/`.WithColumnLetter(...)` or `[ExcelColumnIndex]`, or rename a column. The second of two columns named "Name" can no longer be found as "Name_2", and a column really named "Name_2" is found.
- **Types with nothing to map**: reading rows of a type, or a nested member's type, that maps no members (no public setters or fields, every member `[ExcelIgnore]`, or a registered class map with no `Map(...)` calls) throws an `ExcelMappingException` naming the type, instead of giving an empty object for every row. Map the members in a class map, or ignore the member whose type it is.
- **Changing settings while reading**: a sheet takes the importer's settings (`SkipBlankLines`, `ValidateDataAnnotations`, `Culture`, ...) when it starts reading, so changing them while its rows are read only affects sheets read afterwards. Before 4.0, a change applied to the following rows.
- **Constructor exceptions**: an exception thrown by a constructor called during mapping (for records and constructor-only types) is wrapped in an `ExcelMappingException` with the row, with the original as `InnerException`.
- **Fallback values**: `WithFallback` (was `WithValueFallback`), `WithEmptyFallback` and `WithInvalidFallback` take a value of the member's type, where they took an `object`. A value of the wrong type, such as `WithEmptyFallback("none")` on an `int` member, is now a compile error instead of an `InvalidCastException` when rows are read. `WithEmptyFallback(null)` on a member that can't be null throws an `ArgumentNullException`: give the value, such as `WithEmptyFallback(0)`.
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
- `.WithNumberStyle(...)` and `.WithFormatProvider(...)` for number types other than the built-in ones (on the netstandard builds they work for `int`, `decimal`, `double` and the other built-in types)
- `.WithFormats(...)` for `DateOnly` and `TimeOnly`, and `IReadOnlySet<T>` members

## New in 4.0

Worth a look while upgrading:

- [Collect every error in a file](errors.md#collecting-errors) with `ReadRowResults`, and [check columns up front](errors.md#checking-columns-before-reading) with `ValidateHeading`.
- [Find the heading](reading.md#finding-the-heading) of sheets with titles or notes above it with `FindHeading`.
- [Read part of a sheet](reading.md#reading-part-of-a-sheet), or several tables on one sheet, with `sheet.Region("B3:F20")`.
- [Records and constructor-only types](mapping.md#records-and-constructors) map without a parameterless constructor.
- Map by [column letter](mapping.md#choosing-columns), choose a [culture](types.md#culture) for the whole importer, and set the [offset of `DateTimeOffset` values](types.md#offsets).
- `ExcelImporter.ReadRows<T>("file.xlsx")` reads a file in one line.
- Set a member to the [row number](mapping.md#row-numbers), or collect the [columns no other member reads](mapping.md#extra-columns) into a dictionary.
- [Typed converters](mapping.md#converters) get the cell's value as Excel stores it and where the cell is, without boxing.
- Mapping works in [trimmed apps](trimming.md), and the package includes XML documentation for IntelliSense.
- Mapping is faster and allocates much less: see the [changelog](../CHANGELOG.md#performance).
