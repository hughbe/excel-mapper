# CLAUDE.md

Guidance for working on ExcelMapper (NuGet package `ExcelDataReader.Mapping`), a library that maps Excel and CSV rows to .NET objects on top of ExcelDataReader.

## Layout

- `src/` - the library. Targets `netstandard2.0`, `netstandard2.1` and `net8.0`.
- `tests/` - xUnit tests (`net8.0`). See [Tests](#tests).
- `benchmarks/` - BenchmarkDotNet benchmarks. See [Performance](#performance).
- `docs/` - user documentation, linked from the README's documentation index. `docs/Architecture.md` explains the map/reader/pipeline design.

## Build and test

```bash
dotnet build src -c Release          # builds all three target frameworks
dotnet test tests                     # runs the whole suite
dotnet test tests --filter "FullyQualifiedName~ExcelRangeTests"
./coverage.sh                         # coverage report in tests/coverage
```

Always build `src` in Release before finishing: code must compile for `netstandard2.0`, which lacks newer APIs (for example `init` accessors, `Array.Fill`, `Enum.TryParse(Type, ...)`, span-based `int.TryParse`). Use `#if NET8_0_OR_GREATER` (or similar) where a newer API is needed, and keep behaviour the same on every target. The build should have no warnings.

## Code conventions

- Public types and members have XML doc comments.
- Validate arguments with `ThrowHelpers` (`ThrowIfNull`, `ThrowIfNegative`, ...) so parameter names match the API.
- Problems with a file or a mapping throw `ExcelMappingException`. Create it with the sheet, row and column where known so messages include the cell address (`at cell B11`).
- Nullable reference types are enabled.
- Attributes live in `src/` as `Excel<Name>Attribute.cs`; mappers in `src/Mappers`, readers in `src/Readers`, factories in `src/Factories`, fallbacks in `src/Fallbacks`, internal helpers in `src/Utilities`.

## Tests

Every change comes with tests, structured like the existing ones. There are two layers:

1. **Unit tests in `tests/ExcelMapper/`, one file per public type, mirroring `src/`.** A new public type `src/Foo/BarBaz.cs` gets `tests/ExcelMapper/Foo/BarBazTests.cs` with class `BarBazTests`. For example:
   - `src/ExcelColumnLetterAttribute.cs` → `tests/ExcelMapper/ExcelColumnLetterAttributeTests.cs`
   - `src/Mappers/ObjectMapper.cs` → `tests/ExcelMapper/Mappers/ObjectMapperTests.cs` (namespace `ExcelMapper.Mappers.Tests`)
   - `src/Abstractions/CellMapperResult.cs` → `tests/ExcelMapper/Abstractions/CellMapperResultTests.cs`

   New public members on an existing type are tested in that type's existing test file (for example a new extension method goes in `IOneToOneMapExtensionsTests`, a new mapper behaviour in that mapper's tests).

   Internal types (no `InternalsVisibleTo`) are tested through the public API.

2. **Mapping tests in `tests/Maps/`** (and `tests/Fallbacks/`, `tests/Transformers/`) read real data end to end. Name them `Map<Feature>Tests` and use the pattern `ReadRows_<Scenario>_ReturnsExpected` / `ReadRow_<Scenario>_ThrowsExcelMappingException`.

Test naming in unit tests follows the existing files:

- `Ctor_<Arguments>` for valid constructor calls, for example `Ctor_String`, `Ctor_Int`.
- `Ctor_Null<Arg>_ThrowsArgumentNullException`, `Ctor_Invalid<Arg>_ThrowsArgumentException` - and pass the parameter name: `Assert.Throws<ArgumentException>("letter", () => ...)`.
- `<Property>_GetDefault_Returns<Value>`, `<Property>_Set_GetReturnsExpected` (set, then "Set same"), `<Property>_SetInvalid_Throws<Exception>`.
- `<Method>_<Condition>_<Result>`, for example `Map_ValidStringValue_ReturnsSuccess`, `Map_InvalidStringValue_ReturnsInvalid`, `WithColumnLetter_InvalidColumnLetter_ThrowsArgumentException`.
- Use `[Theory]` with `[InlineData]`, or `[MemberData(nameof(X_TestData))]` with a `public static IEnumerable<object[]> X_TestData()` method, for multiple cases.

Test helpers:

- `Helpers.GetImporter("File.xlsx")` opens a file from `tests/Resources`.
- `Helpers.GetCsvImporter("A,B\n1,2\n")` creates an importer for in-memory CSV (every cell is text).
- `XlsxBuilder.Create(["Header"], [2.5], [new DateTime(2024, 3, 4)])` builds an in-memory .xlsx with typed cells (numbers, booleans, dates, times, durations, `Formatted`, `Formula`, `Error`), so tests don't need new binary fixtures.
- Mapper unit tests use a private `MockExcelDataReader` with `GetValueAction` to supply typed cell values; copy it from a neighbouring mapper test file.

## Performance

Performance matters: the mapping path runs for every cell of every row. For any change to `src` that can affect it:

- Compare against `main`, preferably in-process (interleaved runs, best of several, `GC.GetAllocatedBytesForCurrentThread()` for allocations) or with the `Mapping` benchmark: `cd benchmarks && dotnet run -c Release -- --filter '*Mapping*' --job short`. BenchmarkDotNet timings on Apple Silicon can vary ~2x between runs; allocations are reliable, so compare those, and repeat timings.
- Report the numbers in the PR description.

Patterns to keep:

- Don't throw and catch exceptions for invalid cells: mappers use `TryParse` and `CellMapperResult.Invalid(Func<Exception>)` (or `MapperUtilities.InvalidFromParse`) so the exception is only created if it is read.
- A lambda that captures locals allocates its closure when the locals' scope is entered, even if the lambda isn't used. Create such lambdas in a separate static method that takes the values as parameters.
- Use indexed `for` loops over `IList<T>` in per-cell code; `foreach` over an interface allocates an enumerator.
- Cache per-sheet state in `PerSheetCache<T>` (not a strongly-keyed dictionary), and keep maps safe to share between importers and threads.

## Documentation

Update the docs in the same change as the code:

- `README.md` - overview, quick start and the documentation index. It is also the NuGet package page, so link to `docs/` pages with absolute GitHub URLs.
- `docs/*.md` - the topic pages (reading, mapping, types, errors, performance, API reference). Add new attributes and fluent methods to `docs/api.md` as well as the relevant topic page.
- Note behaviour changes, and mark breaking changes; bump `PackageVersion` in `src/ExcelMapper.csproj` for releases (major version for breaking changes).

## Pull requests

- One feature or fix per PR, branched from `main`. Avoid stacking PRs on each other's branches: a PR merged into another feature branch doesn't reach `main`.
- CI (`.github/workflows/dotnetcore.yml`) runs build and tests for PRs into `main` only.
- PRs are squash-merged.
