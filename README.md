# ExcelMapper

A library that reads a row of an Excel sheet and maps it to an object. A flexible and extensible fluent mapping system allows you to customize the way the row is mapped to an object.

![.NET Core](https://github.com/hughbe/excel-mapper/workflows/.NET%20Core/badge.svg)

## Basic Mapping

ExcelMapper will go through each public property or field and attempt to map the value of the cell in the column with the name of the member. If the column cannot be found or mapped, an exception will be thrown.

| Name          | Location         | Attendance | Date       | Link                    | Revenue | Successful | Cause   |
|---------------|------------------|------------|------------|-------------------------|---------|------------|---------|
| Pub Quiz      | The Blue Anchor  | 20         | 18/07/2017 | http://eventbrite.com   | 100.2   | TRUE       | Charity |
| Live Music    | The Raven        | 15         | 17/07/2017 | http://ticketmaster.com | 105.6   | FALSE      | Profit  |
| Live Football | The Rutland Arms | 45         | 16/07/2017 | http://facebook.com     | 263.9   | TRUE       | Profit  |


```cs
public enum EventCause
{
    Profit,
    Charity
}

public class Event
{
    public string Name { get; set; }
    public string Location { get; set; }
    public int Attendance { get; set; }
    public DateTime Date { get; set; }
    public Uri Link { get; set; }
    public EventCause Cause { get; set; }
}

// ...

using var stream = File.OpenRead("Pub Events.xlsx");
using var importer = new ExcelImporter(stream);

ExcelSheet sheet = importer.ReadSheet();
Event[] events = sheet.ReadRows<Event>().ToArray();
Console.WriteLine(events[0].Name); // Pub Quiz
Console.WriteLine(events[1].Name); // Live Music
Console.WriteLine(events[2].Name); // Live Football
```

You can skip blank lines by setting `importer.Configuration.SkipBlankLines = true`. This is off by default for performance reasons.

## Attribute Mapping

### Custom Column Names
ExcelMapper supports specifying a custom column name with the `ExcelColumnName` attribute. This is useful for cases where we want to map a column name that is different from the property name in the data structure, for example the column name contains whitespace or the column name contains characters that can't be represented in C#.

| Full Name      | #Age |
|----------------|------|
| Donald Trump   | 73   |
| Barack Obama   | 58   |

```cs
public class President
{
    [ExcelColumnName("Full Name")]
    public string Name { get; set; }

    [ExcelColumnName("#Age")]
    public int Age { get; set; }
}


// ...

using var stream = File.OpenRead("Presidents.xlsx");
using var importer = new ExcelImporter(stream);

ExcelSheet sheet = importer.ReadSheet();
President[] president = sheet.ReadRows<President>().ToArray();
Console.WriteLine(president[0].Name); // Donald Trump
Console.WriteLine(president[0].Age); // 73
Console.WriteLine(president[1].Name); // Barack Obama
Console.WriteLine(president[1].Age); // 58
```

### Ignoring Properties
ExcelMapper supports ignoring properties when deserializing with the `ExcelIgnoreAttribute` attribute. This is useful for cases where the the data structure is recursive, or the property is missing data and you don't want to create a custom mapper and call `MakeOptional`.

| Name           | Age |
|----------------|-----|
| Donald Trump   | 73  |
| Barack Obama   | 58  |

```cs
public class President
{
    public string Name { get; set; }

    // If you want to ignore a property that exists for any reason.
    [ExcelIgnore]
    public int Age { get; set; }

    // If you want to ignore a property that doesn't exist for any reason.
    [ExcelIgnore]
    public object NoSuchColumn { get; set; }

    // If the data structure is recursive
    [ExcelIgnore]
    public President PreviousPresident { get; set; }
}


// ...

using var stream = File.OpenRead("Presidents.xlsx");
using var importer = new ExcelImporter(stream);

ExcelSheet sheet = importer.ReadSheet();
President[] president = sheet.ReadRows<President>().ToArray();
Console.WriteLine(president[0].Name); // Donald Trump
Console.WriteLine(president[0].Age); // 0
Console.WriteLine(president[0].NoSuchColumn); // null
Console.WriteLine(president[0].PreviousPresident); // null
Console.WriteLine(president[1].Name); // Barack Obama
Console.WriteLine(president[1].Age); // 0
Console.WriteLine(president[1].NoSuchColumn); // null
Console.WriteLine(president[1].PreviousPresident); // null
```

## Defining a custom mapping

ExcelMapper allows customizing the class map used when mapping a row in an Excel sheet to an object.

Custom class maps allow you to change the column used to map a cell or make that column optional. You can also change the default map of the class, adding custom conversions.

| Name           | Marrital Status | Number of Children | Approval Rating (%) | Date of Birth |
|----------------|-----------------|--------------------|---------------------|---------------|
| Donald Trump   | Twice Married   | 5                  | 60%                 | 1946-06-14    |
| Barack Obama   | Married         | 2                  | 50                  | 04/08/1961    |
| Ronald Reagan  | Divorced        | 5                  | 75                  | 06/03/1911    |
| James Buchanan | Single          | 0                  | 100                 | 1791-04-23    |



```cs
public enum MaritalStatus
{
    Married,
    Divorced,
    Single
}

public class President
{
    public string Name { get; set; }
    public MaritalStatus MarritalStatus { get; set; }
    public int NumberOfChildren { get; set; }
    public float ApprovalRating { get; set; }
    public DateTime DateOfBirth { get; set; }

    public int NumberOfTerms { get; set; }
}

public class PresidentClassMap : ExcelClassMap<President>
{
    public PresidentClassMap()
    {
        // Simple mapping.
        Map(president => president.Name);

        // Map invalid unknown value to a known value.
        Map(president => president.MarritalStatus)
            .WithColumnName("Marrital Status")
            .WithMapping(new Dictionary<string, MaritalStatus>
            {
                { "Twice Married", MaritalStatus.Married    }
            });

        // Read from a column index.
        Map(president => president.NumberOfChildren)
            .WithColumnIndex(2);

        // Read with a custom converter delegate.
        Map(president => president.ApprovalRating)
            .WithConverter(value => int.Parse(value) / 100f)
            .WithColumnName("Approval Rating (%)");

        // Read a date with one or more formats.
        Map(president => president.DateOfBirth)
            .WithColumnName("Date of Birth")
            .WithDateFormats("yyyy-MM-dd", "g");

        // Read a missing column.
        Map(president => president.NumberOfTerms)
            .MakeOptional();
    }
}

// ...

using (var stream = File.OpenRead("Presidents.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    // You can register class maps by type.
    importer.Configuration.RegisterClassMap<PresidentClassMap>();

    // Or by namespace.
    importer.RegisterMapperClassesByNamespace("My.Namespace");

    ExcelSheet sheet = importer.ReadSheet();
    President[] president = sheet.ReadRows<President>().ToArray();
    Console.WriteLine(president[0].Name); // Donald Trump
    Console.WriteLine(president[1].Name); // Barack Obama
    Console.WriteLine(president[2].Name); // Ronald Reagan
    Console.WriteLine(president[2].Name); // James Buchanan
}
```

## Nullables and fallbacks

ExcelMapper will set nullable values to `null` if the value of the cell is empty. By default an exception will be thrown if the value of the cell is invalid and cannot be mapped to the type of the property or field.

You can customize the fallback to use a fixed value if the value of a cell is empty or the value of the cell cannot be mapped.

| Name           | Marrital Status | Number of Children | Date of Birth |
|----------------|-----------------|--------------------|---------------|
| Donald Trump   | Twice Married   | invalid            | invalid       |
| Barack Obama   |                 |                    |               |
| Ronald Reagan  | Divorced        | 5                  | 06/03/1911    |
| James Buchanan | Single          | 0                  | 1791-04-23    |


```cs
public enum MaritalStatus
{
    Married,
    Single,
    Invalid,
    Unknown
}

public class President
{
    public string Name { get; set; }
    public MaritalStatus MarritalStatus { get; set; }
    public int? NumberOfChildren { get; set; }
    public DateTime? DateOfBirth { get; set; }
}

public class PresidentClassMap : ExcelClassMap<President>
{
    public PresidentClassMap()
    {
        Map(president => president.Name);

        Map(president => president.MarritalStatus)
            .WithColumnName("Marrital Status")
            .WithEmptyFallback(MaritalStatus.Unknown)
            .WithInvalidFallback(MaritalStatus.Invalid);

        Map(president => president.NumberOfChildren)
            .WithColumnIndex(2)
            .WithInvalidFallback(-1);

        Map(president => president.DateOfBirth)
            .WithColumnName("Date of Birth")
            .WithDateFormats("yyyy-MM-dd", "g")
            .WithInvalidFallback(null);
    }
}

// ...


using (var stream = File.OpenRead("Presidents.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    // You can register class maps by type.
    importer.Configuration.RegisterClassMap<PresidentClassMap>();

    // Or by namespace.
    importer.RegisterMapperClassesByNamespace("My.Namespace");

    ExcelSheet sheet = importer.ReadSheet();
    President[] president = sheet.ReadRows<President>().ToArray();
    Console.WriteLine(president[0].Name); // Donald Trump
    Console.WriteLine(president[1].Name); // Barack Obama
    Console.WriteLine(president[2].Name); // Ronald Reagan
    Console.WriteLine(president[2].Name); // James Buchanan
}
```

## Mapping enums

ExcelMapper supports mapping string values to an enum. By default this is case sensitive matching the behaviour of the .NET Framework, but this case be overriden in a custom map.

| Name           | Marrital Status | Number of Children | Date of Birth |
|----------------|-----------------|--------------------|---------------|
| Donald Trump   | Married         | invalid            | invalid       |
| Barack Obama   |                 |                    |               |
| Ronald Reagan  | diVorCED        | 5                  | 06/03/1911    |
| James Buchanan | Single          | 0                  | 1791-04-23    |


```cs
public enum MaritalStatus
{
    Married,
    Single,
    Invalid,
    Unknown
}

public class President
{
    public string Name { get; set; }
    public MaritalStatus MarritalStatus { get; set; }
}

public class PresidentClassMap : ExcelClassMap<President>
{
    public PresidentClassMap()
    {
        Map(president => president.Name);

        Map(president => president.MarritalStatus, ignoreCase: true)
            .WithColumnName("Marrital Status")
            .WithEmptyFallback(MaritalStatus.Unknown)
            .WithInvalidFallback(MaritalStatus.Invalid);
    }
}

// ...


using (var stream = File.OpenRead("Presidents.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    // You can register class maps by type.
    importer.Configuration.RegisterClassMap<PresidentClassMap>();

    // Or by namespace.
    importer.RegisterMapperClassesByNamespace("My.Namespace");

    ExcelSheet sheet = importer.ReadSheet();
    President[] president = sheet.ReadRows<President>().ToArray();
    Console.WriteLine(president[0].Name); // Donald Trump
    Console.WriteLine(president[1].Name); // Barack Obama
    Console.WriteLine(president[2].Name); // Ronald Reagan
    Console.WriteLine(president[2].Name); // James Buchanan
}
```

## Mapping enumerables

ExcelMapper supports mapping enumerables and lists. If no column names or column indices are supplied then the value of the cell will be split with the `','` separator.

| Name         | Children Names | First Election | Second Election | First Inauguration | Second Inauguration |
|--------------|----------------|----------------|-----------------|--------------------|---------------------|
| Barack Obama | Malia,Sasha    | 04/11/2008     | 06/11/2012      | 2009-01-20         | 20/01/2013          |

```cs
public class President
{
    public string Name { get; set; }
    public IEnumerable<string> ChildrenNames { get; set; }
    public IEnumerable<DateTime> Elections { get; set; }
    public IEnumerable<DateTime> Inaugurations { get; set; }
}

public class PresidentClassMap : ExcelClassMap<President>
{
    public PresidentClassMap()
    {
        Map(president => president.Name);

        // Default: splits with ",".
        Map(president => president.ChildrenNames)
            .WithColumnName("Children Names");

        // Reads the values of multiple columns in order given.
        Map(president => president.Elections)
            .WithColumnNames("First Election", "Second Election");

        // Reads the values of multiple columns in order given.
        Map(president => president.Inaugurations)
            .WithColumnIndices(4, 5)
            .WithElementMap(m => m
                .WithDateFormats("yyyy-MM-dd", "g")
            );
    }
}

// ...


using (var stream = File.OpenRead("Presidents.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    // You can register class maps by type.
    importer.Configuration.RegisterClassMap<PresidentClassMap>();

    // Or by namespace.
    importer.RegisterMapperClassesByNamespace("My.Namespace");

    ExcelSheet sheet = importer.ReadSheet();
    President[] president = sheet.ReadRows<President>().ToArray();
    Console.WriteLine(president[0].Name); // Barack Obama
    Console.WriteLine(president[0].ChildrenNames); // [Malia, Sasha]
    Console.WriteLine(president[0].Elections); // [2008-11-04, 2012-11-06]
    Console.WriteLine(president[0].Inaugurations); // [2009-01-20, 2013-01-20]
}
```

## Mapping nested objects

ExcelMapper supports mapping nested objects. If no column names or class map is supplied, then the nested object is automatically mapped based on the names of it's public properties and fields.


| Name         | First Elected | Electoral College Votes |
|--------------|---------------|-------------------------|
| Barack Obama | 04/11/2008    | 365                     |

```cs
public class Election
{
    public DateTime Date { get; set; }
    public int ElectoralCollegeVotes { get; set; }
}

public class President
{
    public string Name { get; set; }
    public Election ElectionInformation { get; set; }
}

public class PresidentClassMap : ExcelClassMap<President>
{
    public PresidentClassMap()
    {
        Map(president => president.ElectionInformation.Date)
            .WithColumnName("First Elected");

        Map(president => president.ElectionInformation.ElectoralCollegeVotes)
            .WithColumnName("Electoral College Votes");
    }
}

// ...

using (var stream = File.OpenRead("Presidents.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    // You can register class maps by type.
    importer.Configuration.RegisterClassMap<PresidentClassMap>();

    // Or by namespace.
    importer.RegisterMapperClassesByNamespace("My.Namespace");

    ExcelSheet sheet = importer.ReadSheet();
    President[] president = sheet.ReadRows<President>().ToArray();
    Console.WriteLine(president[0].Name); // Barack Obama
    Console.WriteLine(president[0].ElectionInformation.Date) // 2008-11-04
    Console.WriteLine(president[0].ElectionInformation.ElectoralCollegeVotes) // 365
}
```

## Mapping Dictionaries and ExpandObject

ExcelMapper supports mapping dictionaries and Expando Object. If no column names are supplied, then the dictionary will contain all columns.

| Name           | Age
|----------------|----|
| Barack Obama   | 58 |
| Michelle Obama | 56 |

### Dictionaries

Note that
```cs
using var stream = File.OpenRead("Presidents.xlsx");
using var importer = new ExcelImporter(stream);

ExcelSheet sheet = importer.ReadSheet();
Dictionary<string, string>[] results = sheet.ReadRows<Dictionary<string, string>>().ToArray();
Console.WriteLine(results[0].Count); // 2: { { name: "Barack Obama" }, { age: "58" } }
Console.WriteLine(results[1].Count); // 2: { { name: "Michelle Obama" }, { age: "56" } }
```

### Field Dictionaries

Note that
```cs
public class DataClass
{
    public Dictionary<string, string> Values { get; set; }
    // Or
    public ExpandoObject Values { get; set; }
}

public class DataClassMap : ExcelClassMap<DataClass>
{
    public DataClass()
    {
        // Default: reads all column names.
        Map(d => d.Values);

        // Read custom column names.
        Map(d => d.Values)
            .WithColumnNames("Name", "Age");
    }
}

// ...

using var stream = File.OpenRead("Presidents.xlsx");
using var importer = new ExcelImporter(stream);

// You can register class maps by type.
importer.Configuration.RegisterClassMap<DataClassMap>();

// Or by namespace.
importer.RegisterMapperClassesByNamespace("My.Namespace");

ExcelSheet sheet = importer.ReadSheet();
DataClass[] results = sheet.ReadRows<DataClass>().ToArray();
Console.WriteLine(results[0].Values.Count); // 2: { { name: "Barack Obama" }, { age: "58" } }
Console.WriteLine(results[1].Values.Count); // 2: { { name: "Michelle Obama" }, { age: "56" } }
```

# Mapping sheets without headers

By default, ExcelMapper will read the first row of a sheet as a file header. This can be controlled by setting the boolean property `ExcelSheet.HasHeading`.
ExcelMapper will go through each public property or field and attempt to map the value of the cell in the column with the name of the member. If the column cannot be found or mapped, an exception will be thrown.

|               |                  |
|---------------|------------------|
| Pub Quiz      | The Blue Anchor  |
| Live Music    | The Raven        |
| Live Football | The Rutland Arms |

```cs
public class Event
{
    public string Name { get; set; }
    public string Location { get; set; }
}

public class EventClassMap : ExcelClassMap<Event>
{
    public EventClassMap()
    {
        // Read from a column index.
        Map(event => event.Name)
            .WithColumnIndex(0);

        Map(event => event.Location)
            .WithColumnIndex(1);
    }
}

// ...

using (var stream = File.OpenRead("Pub Events.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    importer.Configuration.RegisterClassMap<EventClassMap>();

    ExcelSheet sheet = importer.ReadSheet();
    sheet.HasHeading = false;

    Event[] events = sheet.ReadRows<Event>().ToArray();
    Console.WriteLine(events[0].Name); // Pub Quiz
    Console.WriteLine(events[1].Name); // Live Music
    Console.WriteLine(events[2].Name); // Live Football
}
```

# Reading Sheets

ExcelMapper supports multiple sheets and provides APIs to read particular sheets.

## Enumerating through all sheets

Use the `ReadSheets` method to enumerate through all sheets in a document. Enumeration is reset at the end.

```cs
using (var stream = File.OpenRead("Sheet.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    foreach (ExcelSheet sheet in importer.ReadSheets())
    {
        // Do something with the sheet.
    }
}
```

## Getting the next sheet

Use the `ReadSheet()` method to read the next sheet. This will throw if there are no more sheets. Use the `TryReadSheet(out ExcelSheet sheet)` method to avoid throwing behaviour.

```cs
using (var stream = File.OpenRead("Sheet.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    // Either:
    ExcelSheet sheet = importer.ReadSheet();

    // Or:
    bool success = importer.TryReadSheet(out ExcelSheet sheet);
}
```

## Getting a sheet by name

Use the `ReadSheet(string sheetName)` method to read a particular sheet by name. This will throw if the sheet is not found. Use the `TryReadSheet(string sheetName, out ExcelSheet sheet)` method to avoid throwing behaviour.

```cs
using (var stream = File.OpenRead("Sheet.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    // Either:
    ExcelSheet sheet = importer.ReadSheet("Sheet Name");

    // Or:
    bool success = importer.TryReadSheet("Sheet Name", out ExcelSheet sheet);
}
```

## Getting a sheet by index

Use the `ReadSheet(int sheetIndex)` method to read a particular sheet at the given zero-based index. This will throw if the sheet is not found. Use the `TryReadSheet(int sheetIndex, out ExcelSheet sheet)` method to avoid throwing behaviour.

```cs
using (var stream = File.OpenRead("Sheet.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    // Either:
    ExcelSheet sheet = importer.ReadSheet(0);

    // Or:
    bool success = importer.TryReadSheet(0, out ExcelSheet sheet);
}
```

# Mapping sheets where the header is not the first row

ExcelMapper supports reading headers which are not in the first row. Set the property `ExcelSheet.HeadingIndex` to the zero-based index of the header. All data preceding the header will be skipped and will not be mapped.

|               |                  |
|---------------|------------------|
|               |                  |
|               |                  |
| Name          | Location         |
| Pub Quiz      | The Blue Anchor  |
| Live Music    | The Raven        |
| Live Football | The Rutland Arms |

```cs
public class Event
{
    public string Name { get; set; }
    public string Location { get; set; }
}

// ...

using (var stream = File.OpenRead("Pub Events.xlsx"))
using (var importer = new ExcelImporter(stream))
{
    ExcelSheet sheet = importer.ReadSheet();
    sheet.HeadingIndex = 3;

    Event[] events = sheet.ReadRows<Event>().ToArray();
    Console.WriteLine(events[0].Name); // Pub Quiz
    Console.WriteLine(events[1].Name); // Live Music
    Console.WriteLine(events[2].Name); // Live Football
}
```
