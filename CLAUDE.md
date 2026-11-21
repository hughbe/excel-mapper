# CLAUDE.md

Guidance for working on ExcelMapper (NuGet package `ExcelDataReader.Mapping`), a library that maps Excel and CSV rows to .NET objects on top of ExcelDataReader.

## Layout

- `src/` - the library. Targets `netstandard2.0`, `netstandard2.1` and `net8.0`.
- `tests/` - xUnit tests (`net8.0`). See [Tests](#tests). `tests/TrimmingTests` is a separate app that CI publishes trimmed and runs.
- `benchmarks/` - BenchmarkDotNet benchmarks. See [Performance](#performance).
- `docs/` - user documentation, linked from the README's documentation index. `docs/Architecture.md` explains the map/reader/pipeline design.

## Build and test

```bash
dotnet build src -c Release          # builds all three target frameworks
dotnet test tests                     # runs the whole suite against the net8.0 build
dotnet test tests -p:LibraryTargetFramework=netstandard2.0   # ...and against the netstandard builds
dotnet test tests -p:LibraryTargetFramework=netstandard2.1
dotnet test tests --filter "FullyQualifiedName~ExcelRangeTests"
dotnet pack src -c Release            # also runs package validation (see below)
./coverage.sh                         # coverage report in tests/coverage
```

Always build `src` in Release before finishing: code must compile for `netstandard2.0`, which lacks newer APIs (for example `init` accessors, `Array.Fill`, `Enum.TryParse(Type, ...)`, span-based `int.TryParse`). Use `#if NET8_0_OR_GREATER` (or similar) where a newer API is needed, and keep behaviour the same on every target. The build should have no warnings.

The netstandard builds are used by .NET Framework and older .NET apps and take different code paths, so run the tests against them too (CI does). Tests of APIs that only exist in the net8.0 build (for example `DateOnly`, `INumberBase<T>` types or `WithNumberStyle`) are wrapped in `#if LIBRARY_NET8`; wrap only the net8-only parts so the rest still runs against netstandard. Note any feature that differs between builds in the changelog's platform notes and `docs/migrating-to-4.0.md`.

## Code conventions

- The net8.0 build runs the trimming and native AOT analyzers (`IsTrimmable`, `IsAotCompatible`) and fails on their warnings. Annotate new reflection: put `[DynamicallyAccessedMembers]` on `Type` parameters and generic parameters where the caller knows the type (see `TrimmingConstants`), and `[RequiresUnreferencedCode]`/`[RequiresDynamicCode]` on code that discovers types at run time, carried up to the public API. Don't suppress warnings unless the justification is true. Changes to what mapping needs at run time should be checked with the trimmed app in `tests/TrimmingTests` (see [docs/trimming.md](docs/trimming.md)).

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
- `CHANGELOG.md` - add every user-visible change to the top (unreleased) section under Breaking changes, Added, Changed, Fixed or Performance, with a link to the PR.
- Breaking changes (including behaviour changes, not only API changes) also go in the migration guide for that major version, for example `docs/migrating-to-4.0.md`: who is affected, before/now, and what to do.

## Releases and compatibility

- `PackageVersion` in `src/ExcelMapper.csproj` follows semantic versioning: a major version for breaking changes.
- Package validation (`EnablePackageValidation`) compares the package with `PackageValidationBaselineVersion` (the last release on NuGet) on every `dotnet pack`, including in CI, and fails on unintended API breaks. Intentional breaks are listed in `src/CompatibilitySuppressions.xml` with a comment explaining why; regenerate it with `dotnet pack src -c Release -p:GenerateCompatibilitySuppressionFile=true`. After publishing a release, move the baseline to it and remove the suppressions it covered.
- Package validation only sees the API: behaviour changes still need the changelog and migration guide.

## Pull requests

- One feature or fix per PR, branched from `main`. Avoid stacking PRs on each other's branches: a PR merged into another feature branch doesn't reach `main`.
- CI (`.github/workflows/dotnetcore.yml`) runs for PRs into `main` only: build, tests with coverage, pack with package validation, the tests against the netstandard builds, and the trimmed test app.
- PRs are squash-merged.
