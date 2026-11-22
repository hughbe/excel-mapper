# ExcelMapper

Map the rows of Excel workbooks and CSV files to strongly-typed .NET objects, by convention, with attributes or with a fluent API.

![.NET Core](https://github.com/hughbe/excel-mapper/workflows/.NET%20Core/badge.svg)
[![Nuget](https://img.shields.io/nuget/v/ExcelDataReader.Mapping)](https://www.nuget.org/packages/ExcelDataReader.Mapping/)

Built on [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader). Supports .NET 8 and later, and .NET Framework and other platforms through .NET Standard 2.0 and 2.1.

## Features

- **Automatic mapping** - Columns map to properties and fields by name, with no configuration
- **Attributes and a fluent API** - Choose columns by name, letter, index or pattern, and control how each value is read
- **Rich types** - Numbers, dates and times, enums, `Guid`, `Uri`, nullable types, collections, dictionaries, nested objects, records and `IParsable<T>` types
- **Real-world sheets** - Find a heading below titles and notes, read part of a sheet or several tables on one sheet, and set row numbers and unmapped columns on your objects
- **Clear errors** - Messages name the cell (`at cell B11`), and every invalid cell in a file can be collected instead of stopping at the first
- **Extensible** - Typed converters, transformers, fallbacks and column selectors
- **Fast** - Rows stream lazily and map with compiled, typed setters, without boxing; trimmed apps are supported

## Installation

```bash
dotnet add package ExcelDataReader.Mapping
```

## Quick Start

| Name          | Department  | Salary |
|---------------|-------------|--------|
| Alice Johnson | Engineering | 95000  |
| Bob Smith     | Marketing   | 78000  |

```csharp
using ExcelMapper;

public class Employee
{
    public string Name { get; set; }
    public string Department { get; set; }
    public decimal Salary { get; set; }
}

var employees = ExcelImporter.ReadRows<Employee>("employees.xlsx");
Console.WriteLine(employees[1].Salary); // 78000
```

`ExcelImporter.ReadRows` reads the first sheet. For other sheets, CSV files and settings, create an importer:

```csharp
using var importer = new ExcelImporter("employees.xlsx");
var sheet = importer.ReadSheet("Staff");
foreach (var employee in sheet.ReadRows<Employee>())
{
    Console.WriteLine(employee.Name);
}
```

## Mapping

With attributes:

```csharp
public record Employee(
    [property: ExcelColumnName("Full Name")] string Name,
    [property: ExcelColumnLetter("C")] decimal Salary,
    [property: ExcelFormats("dd/MM/yyyy")] DateTime HireDate,
    [property: ExcelOptional] string? Notes);
```

Or with a class map:

```csharp
public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(e => e.Name).WithColumnName("Full Name");
        Map(e => e.Department).MakeOptional();
        Map(e => e.Salary).WithColumnLetter("C").WithInvalidFallback(0m);
    }
}

importer.Configuration.RegisterClassMap<EmployeeMap>();
```

## Errors

An invalid cell throws an `ExcelMappingException` that says where it is:

```
Cannot assign "abc" to member "Salary" of type "System.Decimal" in column "Salary" at cell C3 in sheet "Staff".
```

To report every problem in a file at once, read row results:

```csharp
foreach (var result in sheet.ReadRowResults<Employee>())
{
    foreach (var error in result.Errors)
    {
        Console.WriteLine($"{error.CellAddress}: {error.InnerException?.Message}");
    }
}
```

## Documentation

- **[Reading Files](https://github.com/hughbe/excel-mapper/blob/main/docs/reading.md)** - Workbooks and CSV files, sheets, headings, rows, data ranges and regions
- **[Mapping](https://github.com/hughbe/excel-mapper/blob/main/docs/mapping.md)** - Automatic mapping, records, attributes, the fluent API, converters, transformers and fallbacks
- **[Types](https://github.com/hughbe/excel-mapper/blob/main/docs/types.md)** - Supported types, enums, collections, dictionaries and nested objects
- **[Errors and Troubleshooting](https://github.com/hughbe/excel-mapper/blob/main/docs/errors.md)** - Error messages, collecting errors, checking columns and common problems
- **[Performance and Thread Safety](https://github.com/hughbe/excel-mapper/blob/main/docs/performance.md)** - Getting the best performance and using ExcelMapper from several threads
- **[Trimming and Native AOT](https://github.com/hughbe/excel-mapper/blob/main/docs/trimming.md)** - Using ExcelMapper in trimmed apps
- **[API Reference](https://github.com/hughbe/excel-mapper/blob/main/docs/api.md)** - Every public type, attribute and fluent method
- **[Migrating to 4.0](https://github.com/hughbe/excel-mapper/blob/main/docs/migrating-to-4.0.md)** - What to check when upgrading from 3.x, and what's new
- **[Changelog](https://github.com/hughbe/excel-mapper/blob/main/CHANGELOG.md)** - What changed in each version
- **[Architecture](https://github.com/hughbe/excel-mapper/blob/main/docs/Architecture.md)** - How ExcelMapper works, for contributors

## Reader options

Options for ExcelDataReader, such as a password, CSV separators or a fallback encoding, can be passed to the constructor:

```csharp
using ExcelDataReader;

// A password-protected workbook
using var importer = new ExcelImporter("secret.xlsx", ExcelImporterFileType.Excel,
    new ExcelReaderConfiguration { Password = "password" });

// A CSV file that uses semicolons, read as Windows-1252 if it has no byte order mark
using var csvImporter = new ExcelImporter(stream, ExcelImporterFileType.Csv,
    new ExcelReaderConfiguration
    {
        AutodetectSeparators = [';'],
        FallbackEncoding = Encoding.GetEncoding(1252)
    });
```

Code page encodings such as Windows-1252 are registered automatically, so `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)` isn't needed.

## Contributing

Contributions are welcome! See [CLAUDE.md](https://github.com/hughbe/excel-mapper/blob/main/CLAUDE.md) for how the code and tests are organised, and [Architecture](https://github.com/hughbe/excel-mapper/blob/main/docs/Architecture.md) for how it works.

## License

MIT. See the [LICENSE](https://github.com/hughbe/excel-mapper/blob/main/LICENSE) file.
