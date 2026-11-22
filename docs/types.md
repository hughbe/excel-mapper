# Types

The types ExcelMapper maps, and how each one is read from a cell. How to choose columns, convert values yourself and handle empty or invalid cells is in [Mapping](mapping.md). See the [documentation index](../README.md#documentation) for other topics.

## Supported types

| Kind | Types |
|---|---|
| Text | `string`, `char` |
| Numbers | `int`, `long`, `short`, `byte`, `sbyte`, `uint`, `ulong`, `ushort`, `float`, `double`, `decimal`; on .NET 8 also `Half`, `Int128`, `UInt128`, `BigInteger`, `Complex`, `nint`, `nuint` and other `INumberBase<T>` types |
| Dates and times | `DateTime`, `DateTimeOffset`, `TimeSpan`; on .NET 8 also `DateOnly`, `TimeOnly` |
| Other values | `bool`, enums, `Guid`, `Uri`, `Version`, `object`; on .NET 8 any `IParsable<T>` type; any `IConvertible` type |
| Collections | Arrays, lists, sets, queues, stacks, immutable and concurrent collections, and your own (see [Collections](#collections)) |
| Dictionaries | Dictionaries with string keys, `ExpandoObject` and the non-generic dictionaries (see [Dictionaries](#dictionaries)) |
| Objects | Classes, structs and records, mapped member by member, including [nested objects](#nested-objects) |

Nullable versions of every value type are supported: an empty cell is `null`. The ".NET 8" types need the library's `net8.0` build, used by .NET 8 and later; on .NET Framework and other platforms that use the netstandard builds they aren't available.

A cell's value is read as Excel stores it: numbers, dates, times and booleans from Excel files arrive typed, and are converted to the member's type without being turned into text first. CSV cells are always text, and are parsed.

## Numbers

Text is parsed with `NumberStyles.Number`: leading and trailing spaces, a sign, thousands separators and a decimal point, so `" -1,234.5 "` is valid. Floating-point types also accept exponents (`1E5`). Change the style for a member with `[ExcelNumberStyle]` or `.WithNumberStyle(...)`:

| Product  | Price     |
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

A fractional number (such as `2.7`) mapped to an integer type is invalid, as the text `"2.7"` would be: it isn't truncated or rounded. Map to `decimal` or `double`, or give the member an [invalid fallback](mapping.md#empty-and-invalid-cells).

On the netstandard builds, `.WithNumberStyle(...)` and `.WithFormatProvider(...)` aren't available for numbers: use `[ExcelNumberStyle]` and the importer's [culture](#culture).

## Dates and times

Cells that Excel stores as dates, times or durations map directly. Excel stores dates as numbers, the days since 1899-12-30, so a number in a cell without a date format is also read as a date: `45355` is 2024-03-04 for `DateTime`, `DateTimeOffset` and `DateOnly`. For `TimeOnly` a number is a fraction of a day (`0.25` is 06:00), and for `TimeSpan` a number of days (`1.5` is 36 hours). A cell formatted as a time of day maps to `TimeOnly` or `TimeSpan`.

Dates and times written as **text** (every CSV cell, or text cells in a workbook) are parsed in these formats, using the importer's [culture](#culture):

| Type | Accepted by default | Examples (en-GB) |
|------|---------------------|------------------|
| `DateTime`, `DateTimeOffset` | ISO 8601; the culture's date and time (`"G"`), short date and time (`"g"`) and short date (`"d"`) | `2024-03-04`, `2024-03-04T10:30:15Z`, `2024-03-04 10:30`, `04/03/2024 10:30:15`, `04/03/2024` |
| `DateOnly` | the culture's short date (`"d"`); ISO 8601 | `04/03/2024`, `2024-03-04` |
| `TimeOnly` | the culture's short time (`"t"`) and long time (`"T"`); 24-hour times | `10:30`, `10:30:15`, `9:05`, `14:30:15.25` |
| `TimeSpan` | the constant format (`"c"`) | `1:30`, `01:30:00`, `1.02:03:04` |

Text such as `3/4/2024 10:30 AM` matches en-US patterns whether it uses an ordinary space or the narrow no-break space .NET uses on Linux, macOS and recent Windows. A `DateTime` with an offset (`Z`, `+01:00`) is converted to local time, as `DateTime.Parse` does; map to `DateTimeOffset` to keep the offset.

To accept exactly the formats you give, and nothing else, set them with `[ExcelFormats]` or `.WithFormats(...)`:

| Event      | Date       | Start   | Duration |
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
Map(e => e.Date).WithFormats("yyyy-MM-dd", "dd/MM/yyyy");
```

### Offsets

Excel dates have no time zone. Mapped to a `DateTimeOffset`, they get the machine's local offset, so the same file can give different results on different servers. Choose the offset instead:

```csharp
public class Event
{
    [ExcelOffset("+00:00")]          // or .WithOffset(TimeSpan.Zero)
    public DateTimeOffset StartsAt { get; set; }
}
```

This applies to cells Excel stores as dates or numbers. Text with its own offset, parsed with a format such as `"yyyy-MM-ddTHH:mm:sszzz"`, keeps it.

## Culture

Text is parsed, and numbers and dates are converted to text, with the current culture of the thread reading the rows, so results can depend on the machine's settings. Set a culture for the whole importer to make them predictable:

```csharp
importer.Configuration.Culture = CultureInfo.InvariantCulture;   // or CultureInfo.GetCultureInfo("de-DE")
```

The culture is only set while each row is mapped; your code between rows keeps its own. To use a different culture for one member, give it a format provider:

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
        Map(o => o.Price).WithFormatProvider(german);   // numbers: .NET 8 only
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

To ignore case, map the member with `Map(e => e.Status, ignoreCase: true)` in a class map. To map other text, such as codes, use [value mapping](mapping.md#value-mapping).

## Uris

By default only **absolute** URIs (such as `https://example.com`) are valid. Accept relative ones with `[ExcelUri]` or `.WithUriKind(...)`:

```csharp
public class Resource
{
    public Uri Website { get; set; }                   // absolute only

    [ExcelUri(UriKind.Relative)]                       // or .WithUriKind(UriKind.Relative)
    public Uri ApiEndpoint { get; set; }               // "/api/v1/users"

    [ExcelUri(UriKind.RelativeOrAbsolute)]
    public Uri Link { get; set; }                      // either
}
```

A hyperlink cell maps from its displayed text: ExcelDataReader doesn't read hyperlink targets.

## Values as Excel stores them

`object` and `IConvertible` members, and `object` values in dictionaries, `ExpandoObject` and non-generic collections such as `ArrayList`, receive the value as Excel stores it: a `double` for numbers, a `DateTime` for dates, a `bool` for booleans, a `TimeSpan` for durations and a `string` for text. With `[ExcelPreserveFormatting]` they receive the displayed text instead. Collections that only hold one type, such as `StringCollection`, convert values to that type.

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

    [ExcelSeparators(';', '|')]             // or .WithSeparators(';', '|'); strings work too
    public List<string> Tags { get; set; }
}
```

An empty cell gives an empty collection, and a trailing separator (`"a,b,"`) gives a trailing empty value, as `string.Split` does.

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
Map(e => e.Certifications)
    .WithColumnNames("Certified 1", "Certified 2")
    .WithElementMap(e => e
        .WithFormats("yyyy-MM-dd")
        .WithInvalidFallback(DateTime.MinValue));
```

Attributes such as `[ExcelFormats]` and `[ExcelConverter]` on a collection member apply to each element.

**Collection types.** Arrays (including multidimensional arrays mapped by index, such as `Map(p => p.Grid[0, 1])`), `List<T>` and the list, collection and enumerable interfaces, `Collection<T>`, `ReadOnlyCollection<T>`, `ObservableCollection<T>`, `LinkedList<T>`, `HashSet<T>`, `SortedSet<T>`, `ISet<T>`, `Queue<T>`, `Stack<T>`, the concurrent collections, `BlockingCollection<T>`, the immutable collections, `FrozenSet<T>` (.NET 8), and the non-generic `ArrayList`, `Queue`, `Stack` and `StringCollection`. Your own collection types work if they have a parameterless constructor and an `Add` method, or a constructor that takes an `IEnumerable<T>`, `ICollection<T>` or `IReadOnlyList<T>`.

To map elements one by one from different columns, map them by index: `Map(p => p.Values[0]).WithColumnName("First")`.

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
Map(e => e.Department.Name).WithColumnName("Department");
Map(e => e.Department.Location);
```

`MapObject(e => e.Department)` returns the nested object's class map, and `.WithClassMap(...)` on it uses or configures another class map. A type that contains itself (such as `Person.Parent`) can't be auto-mapped: mark the member `[ExcelIgnore]`, or map the type with a class map.
