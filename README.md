# ExcelMapper

A powerful, flexible .NET library for mapping Excel spreadsheet data to strongly-typed C# objects. ExcelMapper provides an intuitive fluent API with extensive customization options, robust type conversion, and comprehensive error handling.

![.NET Core](https://github.com/hughbe/excel-mapper/workflows/.NET%20Core/badge.svg)
[![Nuget](https://img.shields.io/nuget/v/ExcelDataReader.Mapping)](https://www.nuget.org/packages/ExcelDataReader.Mapping/)

Built on top of [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) for reliable Excel file parsing.

## Features

- ✨ **Automatic mapping** - Maps properties by convention with zero configuration
- 🎯 **Type-safe fluent API** - Strongly-typed mapping configuration using expressions
- 🔧 **Extensive customization** - Custom converters, transformers, and fallback strategies
- 📊 **Multiple mapping strategies** - One-to-one, many-to-one, collections, dictionaries
- 🏷️ **Attribute-based mapping** - Simple declarative mapping with attributes
- 🔄 **Flexible column selection** - By name, index, regex pattern, or custom predicate; capture row numbers and unmapped columns
- 🛡️ **Robust error handling** - Optional properties, default values, and custom fallbacks
- 🚀 **High performance** - Streaming API with lazy evaluation and caching
- 📦 **Rich type support** - Primitives, enums, DateTime, collections, nested objects, and more

## Quick Start

```csharp
using ExcelMapper;

// Define your model
public class Employee
{
    public string Name { get; set; }
    public string Department { get; set; }
    public decimal Salary { get; set; }
}

// Read Excel data
var employees = ExcelImporter.ReadRows<Employee>("employees.xlsx");
```

That's it! ExcelMapper automatically maps columns to properties by name. `ExcelImporter.ReadRows` reads the first sheet; for more control (other sheets, ranges, settings), create an importer:

```csharp
using var importer = new ExcelImporter("employees.xlsx");
var sheet = importer.ReadSheet();
var employees = sheet.ReadRows<Employee>().ToArray();
```

## Documentation

- **[Reading Files](https://github.com/hughbe/excel-mapper/blob/main/docs/reading.md)** - Opening workbooks and CSV files, choosing sheets, reading rows, headings and data ranges.
- **[Mapping](https://github.com/hughbe/excel-mapper/blob/main/docs/mapping.md)** - How rows are mapped to objects: automatic mapping, records, attributes, the fluent API and value mapping.
- **[Types and Advanced Mapping](https://github.com/hughbe/excel-mapper/blob/main/docs/types.md)** - Supported types, enums, collections, dictionaries, nested objects, custom converters and transformers.
- **[Errors and Troubleshooting](https://github.com/hughbe/excel-mapper/blob/main/docs/errors.md)** - Error messages, collecting errors, checking columns, fallbacks and common problems.
- **[Performance, Thread Safety and Best Practices](https://github.com/hughbe/excel-mapper/blob/main/docs/performance.md)** - Getting the best performance, using ExcelMapper from several threads, and recommended practices.
- **[Trimming and Native AOT](https://github.com/hughbe/excel-mapper/blob/main/docs/trimming.md)** - Using ExcelMapper in trimmed apps, and what is and isn't supported with native AOT.
- **[API Reference](https://github.com/hughbe/excel-mapper/blob/main/docs/api.md)** - The main classes, attributes and fluent methods.
- **[Architecture](https://github.com/hughbe/excel-mapper/blob/main/docs/Architecture.md)** - How ExcelMapper works internally.
- **[Migrating to 4.0](https://github.com/hughbe/excel-mapper/blob/main/docs/migrating-to-4.0.md)** - What to check when upgrading from 3.x.
- **[Changelog](https://github.com/hughbe/excel-mapper/blob/main/CHANGELOG.md)** - What changed in each version.

## Installation

```bash
dotnet add package ExcelDataReader.Mapping
```

### Code page encodings

ExcelDataReader needs code page encodings such as Windows-1252, which aren't available by default on .NET Core and .NET 5+. `ExcelImporter` registers them automatically (`Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)`), so no setup is needed. Registering them yourself as well is harmless.

### Reader options

Options for the underlying ExcelDataReader reader, such as a password, CSV separators or a fallback encoding, can be passed to the constructor:

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

## Basic Usage

### Simple Example

| Name          | Department  | Salary  |
|---------------|-------------|---------|
| Alice Johnson | Engineering | 95000   |
| Bob Smith     | Marketing   | 78000   |

```csharp
using ExcelMapper;

public class Employee
{
    public string Name { get; set; }
    public string Department { get; set; }
    public decimal Salary { get; set; }
}

using var importer = new ExcelImporter("employees.xlsx");
var sheet = importer.ReadSheet();
var employees = sheet.ReadRows<Employee>().ToArray();

Console.WriteLine(employees[0].Name);      // Alice Johnson
Console.WriteLine(employees[1].Salary);    // 78000
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

Built on top of [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) for robust Excel file parsing.
