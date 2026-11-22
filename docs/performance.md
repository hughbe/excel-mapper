# Performance and Thread Safety

What mapping costs, how to keep it fast, and how to use ExcelMapper from several threads. See the [documentation index](../README.md#documentation) for other topics.

## What mapping costs

Most of the work is done once, so mapping a row does little more than read its cells:

- **Once per type and importer:** the class map is created (by auto-mapping, or from your `ExcelClassMap`) and registered. Compiled setters and constructors are cached for the whole process, so later importers reuse them.
- **Once per sheet:** each member's columns are found in the heading, and the members are bound to the sheet.
- **For each row:** each member's cells are read and converted to the member's type, without boxing, and set with a compiled setter. Rows that are structs, and nested structs, are set in place and never boxed.

As a rough guide, a 10-member class read from a CSV file maps at about 1 µs per row on a modern laptop, most of it spent by ExcelDataReader reading the file. The benchmarks in [`benchmarks/`](../benchmarks) measure common scenarios, and compare them with 3.x (see its [README](../benchmarks/README.md)):

```bash
dotnet run -c Release --project benchmarks -- --filter '*.Mapping.*'
```

### Native AOT

With native AOT, read rows with a class map from a generated `ExcelMappingContext` (see [Native AOT apps](trimming.md#native-aot-apps)): its generated code sets members and calls constructors, so rows allocate the same as with the JIT. Mapping 50,000 rows of a 10-member class from CSV takes about 70 ms with native AOT and 45 ms with the JIT. Most of the difference is reading the file: ExcelDataReader itself is about 65% slower with native AOT, which can't optimize from how the code runs as the JIT does.

## Keeping it fast

1. **Stream rows.** `ReadRows<T>()` maps lazily as you enumerate, so a large file doesn't have to fit in memory:

   ```csharp
   foreach (var employee in sheet.ReadRows<Employee>())
   {
       Process(employee);
   }
   ```

   Call `.ToList()` only when you need every row at once.

2. **Give invalid cells a fallback when they're expected.** An invalid cell without a fallback throws an `ExcelMappingException`, which is expensive if many cells are invalid. With a fallback (`[ExcelInvalidValue]`, `.WithInvalidFallback(...)`), nothing is thrown. To collect errors rather than stop at the first, use `ReadRowResults<T>()`: it costs the same as `ReadRows<T>()` for valid rows, and doesn't throw for invalid ones.

3. **Use typed members.** Numbers, dates, strings and enums are converted without boxing. `object` members, `Dictionary<string, object>` and `ExpandoObject` box every value.

4. **Prefer converters for custom conversion.** `WithConverter(value => ...)` and `ICellConverter<T>` convert without boxing. `WithMapping(dictionary)` boxes values, which is fine for a few columns.

5. **Only preserve formatting where you need it.** `[ExcelPreserveFormatting]` formats each cell's value as Excel displays it, which costs more than reading the value.

6. **Reuse importers' configuration where you can.** Auto-mapping a type is cheap, but it's done again for each importer. For many small files, register a shared class map (see below) rather than auto-mapping in every importer.

Settings such as `SkipBlankLines` (which checks every row for empty cells) and `ValidateDataAnnotations` (which runs the validators on every row) add per-row work, so turn them on only if you need them.

## Thread safety

**An `ExcelImporter` and its sheets should be used by one thread at a time.** A sheet keeps its position in the file, and every sheet of an importer shares the importer's reader. To read files in parallel, use one importer for each file:

```csharp
Parallel.ForEach(paths, path =>
{
    var employees = ExcelImporter.ReadRows<Employee>(path);
    Process(employees);
});
```

To process one file's rows in parallel, read them on one thread and hand them out, for example with `Parallel.ForEach(sheet.ReadRows<Employee>(), Process)`. `Parallel.ForEach` takes rows from the enumerator one at a time, so the sheet is only read by one thread. Don't call `ReadRow` or `ReadRows` on the same sheet from several threads yourself.

**Class maps can be shared.** A class map can be registered in many importers, including importers used on different threads at the same time. A class map can't change once rows have been read with it, and maps don't keep any state for the sheets they read, or keep them alive, so one shared map is safe in a long-running service:

```csharp
private static readonly EmployeeMap s_employeeMap = new();

public List<Employee> Import(Stream stream)
{
    using var importer = new ExcelImporter(stream);
    importer.Configuration.RegisterClassMap(s_employeeMap);
    return importer.ReadSheet().ReadRows<Employee>().ToList();
}
```

Don't change a shared class map (for example by calling `Map(...)` on it) while it's being used.

**Converters, transformers, fallbacks and column selectors** can be called from several threads at once when a class map is shared, so they shouldn't change their own state.
