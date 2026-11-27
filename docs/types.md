# Types

The types ExcelMapper maps, and how each one is read from a cell. How to choose columns, convert values yourself and handle empty or invalid cells is in [Mapping](mapping.md). See the [documentation index](../README.md#documentation) for other topics.

## Supported types

| Kind | Types |
|---|---|
| Text | `string`, `char` |
| Numbers | `int`, `long`, `short`, `byte`, `sbyte`, `uint`, `ulong`, `ushort`, `float`, `double`, `decimal`; on .NET 8 also `Half`, `Int128`, `UInt128`, `BigInteger`, `Complex`, `nint`, `nuint` and other `INumberBase<T>` types |
| Dates and times | `DateTime`, `DateTimeOffset`, `TimeSpan`; on .NET 8 also `DateOnly`, `TimeOnly` |
| Other values | `bool`, enums, `Guid`, `Uri`, `Version`, `object`; on .NET 8 any `IParsable<T>` type, such as `IPAddress` (created by its `Parse` method, so it needn't have a parameterless constructor); any `IConvertible` type with a parameterless constructor |
| Collections | Arrays, lists, sets, queues, stacks, immutable and concurrent collections, and your own (see [Collections](#collections)) |
| Dictionaries | Dictionaries keyed by column name, or by keys of another value type converted from the heading (such as `int`, enums or `DateOnly`), `ExpandoObject` and the non-generic dictionaries (see [Dictionaries](#dictionaries)) |
| Objects | Classes, structs and records, mapped member by member, including [nested objects](#nested-objects) |

Nullable versions of every value type are supported: an empty cell is `null`. The ".NET 8" types need the library's `net8.0` build, used by .NET 8 and later; on .NET Framework and other platforms that use the netstandard builds they aren't available.

A cell's value is read as Excel stores it: numbers, dates, times and booleans from Excel files arrive typed, and are converted to the member's type without being turned into text first. CSV cells are always text, and are parsed.

## Numbers

Text is parsed with `NumberStyles.Number`: leading and trailing spaces, a sign, thousands separators and a decimal point, so `" -1,234.5 "` is valid. Floating-point types also accept exponents (`1E5`). Change the style for a member with `[ExcelNumberStyle]` or `.WithNumberStyle(...)`:

| Name     | Price     |
|----------|-----------|
| Widget A | $1,234.56 |
| Widget B | (250.00)  |

```csharp
public class Product
{
    public string Name { get; set; }

    [ExcelNumberStyle(NumberStyles.Currency)]   // currency symbols, and parentheses for negative numbers
    public decimal Price { get; set; }
}
```

The currency symbol and separators are the [culture](#culture)'s, so this example needs an en-US culture. Number styles only affect text: cells Excel stores as numbers aren't parsed.

A fractional number (such as `2.7`) mapped to an integer type is invalid, as the text `"2.7"` would be: it isn't truncated or rounded. Excel calculates with 15 significant digits, so floating-point error past them isn't a fraction: `=1.15*100`, which Excel stores as 114.99999999999999 and shows as 115, maps to 115. Map to `decimal` or `double`, or give the member an [invalid fallback](mapping.md#empty-and-invalid-cells).

On the netstandard builds, `.WithNumberStyle(...)` and `.WithFormatProvider(...)` are available for the built-in number types (`int`, `decimal`, `double` and so on, and their nullable versions).

## Dates and times

Cells that Excel stores as dates, times or durations map directly. Excel stores dates as numbers, the days since 1899-12-30, so a number in a cell without a date format is also read as a date: `45355` is 2024-03-04 for `DateTime`, `DateTimeOffset` and `DateOnly`. For `TimeOnly` a number is a fraction of a day (`0.25` is 06:00), and for `TimeSpan` a number of days (`1.5` is 36 hours). A cell formatted as a time of day maps to `TimeOnly` or `TimeSpan`.

Dates and times written as **text** (every CSV cell, or text cells in a workbook) are parsed in these formats, using the importer's [culture](#culture):

| Type | Accepted by default | Examples (invariant culture) |
|------|---------------------|------------------|
| `DateTime`, `DateTimeOffset` | ISO 8601; the culture's date and time (`"G"`), short date and time (`"g"`) and short date (`"d"`) | `2024-03-04`, `2024-03-04T10:30:15Z`, `2024-03-04 10:30`, `03/04/2024 10:30:15`, `03/04/2024` |
| `DateOnly` | the culture's short date (`"d"`); ISO 8601 | `03/04/2024`, `2024-03-04` |
| `TimeOnly` | the culture's short time (`"t"`) and long time (`"T"`); 24-hour times | `10:30`, `10:30:15`, `9:05`, `14:30:15.25` |
| `TimeSpan` | the constant format (`"c"`) | `1:30`, `01:30:00`, `1.02:03:04` |

Text such as `3/4/2024 10:30 AM` matches en-US patterns whether it uses an ordinary space or the narrow no-break space .NET uses on Linux, macOS and recent Windows. A `DateTime` with an offset (`Z`, `+01:00`) is converted to UTC, with `Kind` `Utc`, rather than to the machine's local time, so every machine reads the same value; text without an offset is left as it is, with `Kind` `Unspecified`. Map to `DateTimeOffset` to keep the offset.

To accept exactly the formats you give, and nothing else, set them with `[ExcelFormats]` or `.WithFormats(...)`:

| Name       | Date       | Start   | Duration |
|------------|------------|---------|----------|
| Conference | 2024-03-15 | 09:30   | 02:30:00 |
| Workshop   | 15/03/2024 | 2:00 PM | 01:15:00 |

```csharp
public class Event
{
    public string Name { get; set; }

    [ExcelFormats("yyyy-MM-dd", "dd/MM/yyyy")]
    public DateTime Date { get; set; }

    [ExcelFormats("HH:mm", "h:mm tt")]
    public TimeOnly Start { get; set; }

    [ExcelFormats(@"hh\:mm\:ss", @"mm\:ss")]
    public TimeSpan Duration { get; set; }
}

// Or with a class map
importer.Configuration.RegisterClassMap<Event>(c => c.Map(e => e.Date).WithFormats("yyyy-MM-dd", "dd/MM/yyyy"));
```

### Offsets

Excel dates have no time zone, and neither does text such as `2024-03-04 10:30`. Mapped to a `DateTimeOffset`, they're UTC by default (`+00:00`), so every machine reads the same value. Choose another offset for a member:

```csharp
public class Event
{
    [ExcelOffset("-08:00")]          // or .WithOffset(TimeSpan.FromHours(-8))
    public DateTimeOffset StartsAt { get; set; }
}
```

This applies to cells Excel stores as dates or numbers, and text without an offset. Text with its own offset (`2024-03-04T10:30:15+01:00`, or a format such as `"yyyy-MM-ddTHH:mm:sszzz"`) keeps it.

## Culture

Text is parsed, and numbers and dates are converted to text, with the invariant culture by default, so results don't depend on the machine's settings: `1.5` is one and a half, and dates are written `03/04/2024 00:00:00` (month first). Heading cells that are numbers or dates are named the same way, so `[ExcelColumnName("2024")]` finds a heading of the number 2024 on every machine. For a file written in another culture, set the importer's culture:

```csharp
importer.Configuration.Culture = CultureInfo.GetCultureInfo("de-DE");   // reads "1,5" and "15.03.2024"
```

To use the culture of the machine reading the file, as 3.x did, set it to `CultureInfo.CurrentCulture`.

The culture is given to the built-in conversions; it isn't set as the thread's culture, so your code, including converters and fallbacks, runs in its own. A converter gets the importer's culture from `context.FormatProvider` (see [Converters](mapping.md#converters)); a `WithConverter(value => ...)` delegate that parses text should pass a culture itself. To use a different culture for one member, give it a format provider:

| Item  | Date       | Price    |
|-------|------------|----------|
| Kabel | 15.03.2024 | 1.234,56 |

```csharp
public class Order
{
    public string Item { get; set; }
    public DateTime Date { get; set; }
    public decimal Price { get; set; }
}

public class OrderMap : ExcelClassMap<Order>
{
    public OrderMap()
    {
        var german = CultureInfo.GetCultureInfo("de-DE");

        Map(o => o.Item);
        Map(o => o.Date).WithFormatProvider(german);
        Map(o => o.Price).WithFormatProvider(german);
    }
}
```

## Booleans

Excel booleans map directly. Text is valid if it's `true` or `false`, `1` or `0`, or `yes` or `no`, ignoring case.

## Enums

Text is matched to the enum's names, case-sensitively, and to the text of `[Description]` attributes on its values (which win when both match):

| Name          | Status             |
|---------------|--------------------|
| Alice Johnson | Full Time Employee |
| Bob Smith     | PartTime           |

```csharp
using System.ComponentModel;

public enum EmploymentStatus
{
    [Description("Full Time Employee")]
    FullTime,

    PartTime,
}

public class Employee
{
    public string Name { get; set; }
    public EmploymentStatus Status { get; set; }
}
```

To ignore case, map the member with `Map(e => e.Status).WithIgnoreCase()` in a class map. To map other text, such as codes, use [value mapping](mapping.md#value-mapping).

Values are parsed as `Enum.Parse` parses them, so a number is the enum value with that number: `1` is `PartTime`. A number that isn't one of the enum's values, such as `99`, is accepted too, and gives a value without a name. Names separated by commas are combined, as for a `[Flags]` enum: `"FullTime, PartTime"` is `FullTime | PartTime`. To accept only the names, list them with [value mapping](mapping.md#value-mapping) and `MappingDictionaryBehavior.Required`, or check the value with `Enum.IsDefined` after reading.

## Uris

By default only **absolute** URIs (such as `https://example.com`) are valid. Accept relative ones with `[ExcelUriKind]` or `.WithUriKind(...)`:

```csharp
public class Resource
{
    public Uri Website { get; set; }                   // absolute only

    [ExcelUriKind(UriKind.Relative)]                       // or .WithUriKind(UriKind.Relative)
    public Uri ApiEndpoint { get; set; }               // "/api/v1/users"

    [ExcelUriKind(UriKind.RelativeOrAbsolute)]
    public Uri Link { get; set; }                      // either
}
```

A hyperlink cell maps from its displayed text: ExcelDataReader doesn't read hyperlink targets.

## Values as Excel stores them

`object` and `IConvertible` members, and `object` values in dictionaries, `ExpandoObject` and non-generic collections such as `ArrayList`, receive the value as Excel stores it: a `double` for numbers, a `DateTime` for dates, a `bool` for booleans, a `TimeSpan` for durations and a `string` for text. With `[ExcelPreserveFormatting]` they receive the displayed text instead. Collections that only hold one type, such as `StringCollection`, convert values to that type.

A formula error, such as `#N/A` or `#DIV/0!`, isn't a value of any type, so it's [invalid](mapping.md#empty-and-invalid-cells) for these members too. A row whose only values are errors isn't blank.

## Collections

A collection member reads one cell split by separators, or several columns.

**One cell, split.** By default a collection reads the column with its name, split by commas:

| Name          | Skills         |
|---------------|----------------|
| Alice Johnson | C#,Python,SQL  |

```csharp
public class Employee
{
    public string Name { get; set; }
    public string[] Skills { get; set; }    // ["C#", "Python", "SQL"]

    [ExcelSeparators(';', '|', Options = StringSplitOptions.RemoveEmptyEntries)]   // or .WithSeparators(';', '|').WithSplitOptions(...)
    public List<string> Tags { get; set; }
}
```

An empty cell gives an empty collection, and a trailing separator (`"a,b,"`) gives a trailing empty value, as `string.Split` does: an empty value uses the element's empty fallback (`null` for strings). `StringSplitOptions.RemoveEmptyEntries` leaves empty values out, and `TrimEntries` (.NET 5 and later) trims each value.

**Several columns**, each giving one element:

```csharp
public class Employee
{
    [ExcelColumnNames("Q1", "Q2", "Q3", "Q4")]           // or .WithColumnNames(...)
    public int[] QuarterlyScores { get; set; }

    [ExcelColumnIndices(5, 6)]                           // or .WithColumnIndices(...)
    public decimal[] Bonuses { get; set; }

    [ExcelColumnsMatching(@"^Review \d+$")]              // or .WithColumnsMatching(regex or predicate)
    public List<string> Reviews { get; set; }
}
```

A [column selector](mapping.md#column-selectors) can choose the columns too.

**Elements** are read like single values of the element type. Configure them with `WithElementMap`, which takes the same value methods as a single member:

```csharp
importer.Configuration.RegisterClassMap<Employee>(c => c
    .Map(e => e.Certifications)
    .WithColumnNames("Certified 1", "Certified 2")
    .WithElementMap(e => e
        .WithFormats("yyyy-MM-dd")
        .WithInvalidFallback(DateTime.MinValue)));
```

Attributes such as `[ExcelFormats]` and `[ExcelConverter]` on a collection member apply to each element.

Elements are read from cells, so they must be values: strings, numbers, dates, enums and the other types above. A collection of objects, such as `List<Address>`, can't be auto-mapped, and throws an `ExcelMappingException` naming the member; the same goes for a dictionary of objects. Give the element type a converter (`[ExcelConverter]` on the member converts each element), map the objects as [nested objects](#nested-objects) instead, or ignore the member with `[ExcelIgnore]`.

**Collection types.** Arrays (including multidimensional arrays mapped by index, such as `Map(p => p.Grid[0, 1])`), `List<T>` and the list, collection and enumerable interfaces, `Collection<T>`, `ReadOnlyCollection<T>`, `ObservableCollection<T>`, `LinkedList<T>`, `HashSet<T>`, `SortedSet<T>`, `ISet<T>`, `Queue<T>`, `Stack<T>`, the concurrent collections, `BlockingCollection<T>`, the immutable collections, `FrozenSet<T>` (.NET 8), and the non-generic `ArrayList`, `Queue`, `Stack` and `StringCollection`. Your own collection types work if they have a parameterless constructor and an `Add` method, or a constructor that takes an `IEnumerable<T>`, `ICollection<T>` or `IReadOnlyList<T>`.

In a class map, `Map(p => p.Tags)` gives a collection map for arrays, `List<T>`, the collection classes in `System.Collections.ObjectModel` and the collection interfaces. For another collection type, such as `HashSet<T>`, use `MapCollection<string>(p => p.Tags)`.

To map elements one by one from different columns, map them by index: `Map(p => p.Values[0]).WithColumnName("First")`. The index (or a dictionary key) can be a constant or a variable, such as a loop's: `for (var i = 0; i < 4; i++) c.Map(p => p.Quarters[i]).WithColumnName($"Q{i + 1}")`. It can't be computed from the row.

## Dictionaries

A dictionary member, or a whole row read as a dictionary, maps every column, keyed by column name:

```csharp
var rows = sheet.ReadRows<Dictionary<string, string>>().ToArray();
Console.WriteLine(rows[0]["Name"]);

foreach (dynamic row in sheet.ReadRows<ExpandoObject>())
{
    Console.WriteLine(row.Name);
}

public class Record
{
    public Dictionary<string, int> Scores { get; set; }   // every column

    [ExcelColumnNames("Maths", "English")]                // or chosen columns
    public Dictionary<string, int> Core { get; set; }
}
```

**Keys of other types.** A dictionary whose keys aren't strings converts each column's heading cell to a key as a cell of the key type would be converted: numbers, enums (by name, `[Description]` or number), dates, `Guid`, `Version`, `Uri` and the other types above. A heading cell that Excel stores as a number or date is converted as itself, so a sheet with a column for each month maps to dates:

| Region | 2024-01-01 | 2024-02-01 |
|--------|------------|------------|
| North  | 1200       | 1350       |

```csharp
public class Sales
{
    public string Region { get; set; }

    [ExcelColumnIndices(1, 2)]
    public Dictionary<DateTime, decimal> Monthly { get; set; }   // { [2024-01-01] = 1200, [2024-02-01] = 1350 }
}
```

Keys are converted once for each sheet, with the importer's [culture](#culture). A heading cell that can't be converted throws an `ExcelMappingException` naming the column. `object` keys are the column names.

Values are read like single values of the value type, and configured with `WithValueMap(v => ...)`. To map only the columns nothing else reads, use [extra columns](mapping.md#extra-columns); to map one value by key, `Map(p => p.Scores["Maths"])`.

**Dictionary types.** `Dictionary<string, TValue>` and the dictionary interfaces, `SortedDictionary`, `SortedList`, `ConcurrentDictionary`, `ReadOnlyDictionary`, the immutable dictionaries, `FrozenDictionary` (.NET 8), `ExpandoObject`, and the non-generic `Hashtable`, `SortedList`, `StringDictionary` and `NameValueCollection`.

## Nested objects

A member whose type is a class, struct or record is mapped from the same row, member by member:

| Name          | Department  | Location |
|---------------|-------------|----------|
| Alice Johnson | Engineering | Seattle  |

```csharp
public class Employee
{
    public string Name { get; set; }
    public DepartmentInfo Department { get; set; }
}

public class DepartmentInfo
{
    [ExcelColumnName("Department")]
    public string Name { get; set; }

    public string Location { get; set; }
}

// Or with a class map
importer.Configuration.RegisterClassMap<Employee>(c =>
{
    c.Map(e => e.Department.Name).WithColumnName("Department");
    c.Map(e => e.Department.Location);
});
```

`MapObject(e => e.Department)` returns the nested object's class map, and `.WithClassMap(...)` on it uses or configures another class map. A type that contains itself (such as `Person.Parent`) can't be auto-mapped: mark the member `[ExcelIgnore]`, or map the type with a class map.
