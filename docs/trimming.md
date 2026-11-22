# Trimming and Native AOT

Using ExcelMapper in apps that are trimmed (`PublishTrimmed`) or compiled with native AOT (`PublishAot`). See the [documentation index](../README.md#documentation) for other topics.

## What to expect

ExcelMapper maps rows by reflection: it finds a type's properties, fields and constructors at run time, and creates generic types and methods for their types. The trimmer can't see all of that, so:

- **The reflection-based APIs are marked** with `[RequiresUnreferencedCode]` and `[RequiresDynamicCode]`: `ReadRows<T>`, `ReadRow<T>`, `TryReadRow<T>`, `ReadRowResults<T>`, `ValidateHeading<T>`, `FindHeading<T>`, the same methods on `ExcelRegion`, `ExcelImporter.ReadRows<T>`, the `ExcelClassMap<T>.Map(...)` methods and `RegisterClassMapsInNamespace`. A trimmed or AOT app gets warnings IL2026 and IL3050 where it calls them, rather than a library that silently stops working.
- **The type you map is kept.** `ReadRows<T>`, `RegisterClassMap<T>` and `ExcelClassMap<T>` tell the trimmer to keep the public constructors, properties and fields of `T`, so classes and records whose members are numbers, strings, dates, enums, lists or dictionaries map in a trimmed app.
- **Types given in attributes are kept**: the types in `[ExcelConverter]`, `[ExcelColumnMatching]`, `[ExcelColumnsMatching]`, `[ExcelEmptyFallback]`, `[ExcelInvalidFallback]` and `[ExcelTransformer]`.
- **Nested types aren't kept automatically.** A member whose type is another class (for example `Customer.Address`) needs that type's members kept by the app, or mapping throws `ExcelMappingException` ("Cannot auto-map type").
- **Native AOT needs a mapping context.** Auto-mapping creates generic code for the types of members it finds at run time, which native AOT can't do unless it has compiled that code. An `ExcelMappingContext` registers the types, and reading rows with its class maps works with native AOT and doesn't warn (see [Native AOT apps](#native-aot-apps)). Without one, class maps registered with `RegisterClassMap` whose members are built-in types (numbers, strings, dates, booleans, GUIDs, nullable values and lists of them) still map, with warnings; auto-mapped members of value types, enums and `IParsable<T>` types throw `NotSupportedException`.

## Trimmed apps

1. Map types whose members are kept: the type passed to `ReadRows<T>` is kept for you. For nested types, keep their members, for example with `[DynamicDependency]`:

   ```csharp
   using System.Diagnostics.CodeAnalysis;

   [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicProperties, typeof(Address))]
   static void KeepMappedTypes()
   {
   }
   ```

   Call it (or put the attribute on a method that is called) before reading rows.

2. Once you've checked the mapped types are kept, suppress the warnings where you call ExcelMapper:

   ```csharp
   using System.Diagnostics.CodeAnalysis;

   [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Mapped types are kept: see KeepMappedTypes.")]
   static List<Customer> ReadCustomers(string path) => ExcelImporter.ReadRows<Customer>(path);
   ```

3. Test the published app: problems with trimming only show up at run time.

The repository's [trimming test app](../tests/TrimmingTests/Program.cs) publishes a trimmed app that maps classes, records, class maps, split lists, dictionaries, enums and nested objects, and runs in CI.

## Native AOT apps

Read rows with class maps from an `ExcelMappingContext`. A context registers the types rows are mapped to (`Row<T>()`) and every type auto-mapping meets in their members, and gives each row type's class map, auto-mapped from its members and attributes as `ReadRows<T>()` does:

```csharp
public sealed class AppMappingContext : ExcelMappingContext
{
    public static AppMappingContext Default { get; } = new();

    public ExcelClassMap<Employee> Employee => GetClassMap<Employee>();

    protected override void RegisterTypes(ExcelTypeRegistrar types)
    {
        types.Row<Employee>();
        types.Type<string>();
        types.Type<decimal>();
        types.Type<int>();
        types.Type<float>();
        types.Type<bool>();
        types.Type<DateTime>();
        types.Array<DateTime>();
    }
}

var employees = sheet.ReadRows(AppMappingContext.Default.Employee).ToList();
```

Every reading method takes a class map: `ReadRows(classMap)`, `ReadRow(classMap)`, `ReadRowResults(classMap)`, `ValidateHeading(classMap)`, `FindHeading(classMap)` and the others, on sheets and regions. They read with the class map you give them rather than looking one up or auto-mapping, so they don't warn.

Register:

- `Row<T>()` for each type rows are mapped to, and `Type<T>()` for the type of each member (and of each type mapped through its own members, such as a nested class or struct).
- `Nullable<T>()` for each nullable value type, such as `int?`.
- `Enum<TEnum>()` for each enum, and `Parsable<T>()` or `ParsableClass<T>()` for each other `IParsable<T>` type.
- `Collection<TCollection, TElement>()`, `Array<TElement>()` or `Dictionary<TDictionary, TKey, TValue>()` for each collection, array or dictionary.

A type that isn't registered still maps with the JIT, but with native AOT throws `NotSupportedException` naming the generic type it needs. The [native AOT test app](../tests/AotTests/Program.cs) shows contexts and what works without one, and runs in CI. Test your published app, as what native AOT is missing only shows up at run time.

Registering class maps by scanning an assembly (`RegisterClassMapsInNamespace`) can't be made trim-safe: register class maps with `RegisterClassMap` instead.
