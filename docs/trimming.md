# Trimming and Native AOT

Using ExcelMapper in apps that are trimmed (`PublishTrimmed`) or compiled with native AOT (`PublishAot`). See the [documentation index](../README.md#documentation) for other topics.

## What to expect

ExcelMapper maps rows by reflection: it finds a type's properties, fields and constructors at run time, and creates generic types and methods for their types. The trimmer can't see all of that, so:

- **The reflection-based APIs are marked** with `[RequiresUnreferencedCode]` and `[RequiresDynamicCode]`: `ReadRows<T>`, `ReadRow<T>`, `TryReadRow<T>`, `ReadRowResults<T>`, `ValidateHeading<T>`, `FindHeading<T>`, the same methods on `ExcelRegion`, `ExcelImporter.ReadRows<T>`, the `ExcelClassMap<T>.Map(...)` methods and `RegisterClassMapsInNamespace`. A trimmed or AOT app gets warnings IL2026 and IL3050 where it calls them, rather than a library that silently stops working.
- **The type you map is kept.** `ReadRows<T>`, `RegisterClassMap<T>` and `ExcelClassMap<T>` tell the trimmer to keep the public constructors, properties and fields of `T`, so classes and records whose members are numbers, strings, dates, enums, lists or dictionaries map in a trimmed app.
- **Types given in attributes are kept**: the types in `[ExcelConverter]`, `[ExcelColumnMatching]`, `[ExcelColumnsMatching]`, `[ExcelEmptyFallback]`, `[ExcelInvalidFallback]` and `[ExcelTransformer]`.
- **Nested types aren't kept automatically.** A member whose type is another class (for example `Customer.Address`) needs that type's members kept by the app, or mapping throws `ExcelMappingException` ("Cannot auto-map type").
- **Native AOT isn't supported yet.** Mapping creates generic types and methods for the member types at run time, which AOT can't guarantee.

## Trimmed apps

1. Map types whose members are kept: the type passed to `ReadRows<T>` is kept for you. For nested types, keep their members, for example with `[DynamicDependency]`:

   ```csharp
   [DynamicDependency(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.PublicProperties, typeof(Address))]
   static void KeepMappedTypes()
   {
   }
   ```

   Call it (or put the attribute on a method that is called) before reading rows.

2. Once you've checked the mapped types are kept, suppress the warnings where you call ExcelMapper:

   ```csharp
   [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Mapped types are kept: see KeepMappedTypes.")]
   static List<Customer> ReadCustomers(string path) => ExcelImporter.ReadRows<Customer>(path);
   ```

3. Test the published app: problems with trimming only show up at run time.

The repository's [trimming test app](../tests/TrimmingTests/Program.cs) publishes a trimmed app that maps classes, records, class maps, split lists, dictionaries, enums and nested objects, and runs in CI.

Registering class maps by scanning an assembly (`RegisterClassMapsInNamespace`) can't be made trim-safe: register class maps with `RegisterClassMap` instead.
