# Performance, Thread Safety and Best Practices

Getting the best performance, using ExcelMapper from several threads, and recommended practices. See the [documentation index](../README.md#documentation) for other topics.

## Performance Tips

1. **Use streaming**: `ReadRows<T>()` uses lazy evaluation - don't materialize unnecessarily
   ```csharp
   // Good - processes one at a time
   foreach (var employee in sheet.ReadRows<Employee>())
   {
       ProcessEmployee(employee);
   }

   // Avoid - loads everything into memory
   var allEmployees = sheet.ReadRows<Employee>().ToList();
   ```

2. **Register maps once**: Class maps are cached per type
   ```csharp
   importer.Configuration.RegisterClassMap<EmployeeMap>();
   ```

3. **Disable blank line skipping**: Off by default for performance
   ```csharp
   importer.Configuration.SkipBlankLines = false;  // Default
   ```

4. **Use column indices for headerless sheets**: Faster than column name lookup
   ```csharp
   Map(e => e.Name).WithColumnIndex(0);  // Faster
   Map(e => e.Name).WithColumnName("Name");  // Requires lookup
   ```

## Thread Safety

**Important:** `ExcelSheet` instances maintain mutable state (current row index) and are **not thread-safe**. Each instance should be used by only one thread at a time.

**Safe concurrent processing:**

```csharp
// Read all rows first
var employees = sheet.ReadRows<Employee>().ToList();

// Then process in parallel
Parallel.ForEach(employees, employee =>
{
    ProcessEmployee(employee);
});
```

**Unsafe concurrent processing:**

```csharp
// DON'T DO THIS - Not thread-safe!
Parallel.ForEach(sheet.ReadRows<Employee>(), employee =>
{
    ProcessEmployee(employee);
});
```

**Multiple sheets:**

```csharp
// Each sheet can be processed independently
using var importer = new ExcelImporter("workbook.xlsx");

foreach (var sheet in importer.ReadSheets())
{
    // Each sheet is independent and can be processed separately
    var data = sheet.ReadRows<Employee>().ToList();
    
    // Now safe to process in parallel
    Parallel.ForEach(data, row => ProcessRow(row));
}
```

**Sharing class maps between importers:**

A class map instance can be registered in many importers, including importers used on different threads at the same time. Maps don't keep sheets alive, so a single shared map is safe in a long-running service:

```csharp
private static readonly EmployeeMap s_employeeMap = new();

public IReadOnlyList<Employee> Import(Stream stream)
{
    using var importer = new ExcelImporter(stream);
    importer.Configuration.RegisterClassMap(s_employeeMap);
    return importer.ReadSheet().ReadRows<Employee>().ToList();
}
```

The importer itself, and each sheet, should still be used by one thread at a time.

## Best Practices

1. **Register class maps once** - Class maps are cached, so register them during application startup
   ```csharp
   importer.Configuration.RegisterClassMap<EmployeeMap>();
   ```

2. **Use streaming for large files** - Avoid `.ToList()` or `.ToArray()` unless necessary
   ```csharp
   // Good
   foreach (var employee in sheet.ReadRows<Employee>())
       ProcessEmployee(employee);
   
   // Avoid if dataset is large
   var allEmployees = sheet.ReadRows<Employee>().ToList();
   ```

3. **Handle missing columns gracefully** - Use `[ExcelOptional]` for non-critical columns
   ```csharp
   [ExcelOptional]
   public string MiddleName { get; set; }
   ```

4. **Provide fallback values** - Make your code resilient to data quality issues
   ```csharp
   [ExcelDefaultValue(0)]
   [ExcelInvalidValue(-1)]
   public int YearsOfService { get; set; }
   ```

5. **Use meaningful error messages** - Custom fallbacks can provide better diagnostics
   ```csharp
   public class SalaryFallback : ICellFallback<decimal>
   {
       public decimal Fallback(CellValue value, CellContext context, Exception? exception)
           => throw new ExcelMappingException($"Invalid salary \"{value.Text}\" at cell {context.CellAddress}.", exception);
   }
   ```

6. **Dispose resources properly** - Always use `using` statements with `ExcelImporter`
   ```csharp
   using var importer = new ExcelImporter("data.xlsx");
   // Work with importer
   // Automatically disposed at end of scope
   ```

7. **Be mindful of thread safety** - Don't share `ExcelSheet` instances across threads
