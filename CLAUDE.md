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

Incremental builds can hide warnings from files that didn't change: use `dotnet build src --no-incremental` to see them all.

The netstandard builds are used by .NET Framework and older .NET apps and take different code paths, so run the tests against them too (CI does). Tests of APIs that only exist in the net8.0 build (for example `DateOnly`, `INumberBase<T>` types or `WithNumberStyle`) are wrapped in `#if LIBRARY_NET8`; wrap only the net8-only parts so the rest still runs against netstandard. Note any feature that differs between builds in the changelog's platform notes and `docs/migrating-to-4.0.md`.

## Code conventions

- The net8.0 build runs the trimming and native AOT analyzers (`IsTrimmable`, `IsAotCompatible`) and fails on their warnings. Annotate new reflection: put `[DynamicallyAccessedMembers]` on `Type` parameters and generic parameters where the caller knows the type (see `TrimmingConstants`), and `[RequiresUnreferencedCode]`/`[RequiresDynamicCode]` on code that discovers types at run time, carried up to the public API. Don't suppress warnings unless the justification is true. Changes to what mapping needs at run time should be checked with the trimmed app in `tests/TrimmingTests` (see [docs/trimming.md](docs/trimming.md)).

- Public types and members have XML doc comments. The documentation ships in the package (IntelliSense), and the build fails on missing or broken doc comments: undocumented public members, wrong `<param>` names, unresolved `cref`s.
- Validate arguments with `ThrowHelpers` (`ThrowIfNull`, `ThrowIfNegative`, ...) so parameter names match the API.
- Problems with a file or a mapping throw `ExcelMappingException`. Create it with the sheet, row and column where known so messages include the cell address (`at cell B11`).
- Nullable reference types are enabled.
- Attributes live in `src/` as `Excel<Name>Attribute.cs`; mappers in `src/Mappers`, readers in `src/Readers`, factories in `src/Factories`, fallbacks in `src/Fallbacks`, internal helpers in `src/Utilities`. Everything in those folders, and in `src/Abstractions`, is internal.
- Values are converted by `ValuePipeline<T>` as `T`, without boxing: mappers implement `ICellMapper<T>` (and `ICellMapper<T?>` for value types, so nullable members share them), fallbacks are `ICellFallback<T>`, and there's no untyped path. A new mapper implements `Convert` and `CreateException` (see `docs/Architecture.md`).
- Generic types created at run time with `MakeGenericType` need a `[DynamicDependency]` on the method that creates them, or the trimmer removes their constructors (check with the trimmed app).
- `Type.GetConstructor(types)` also matches constructors whose parameters the types are assignable to (asking for `ISet<T>` finds a constructor taking `IEnumerable<T>`). Use `CollectionUtilities.GetConstructorTaking` where the parameter type must match exactly.
- Delete code that can't be reached rather than keeping it "just in case": validation that every internal caller already guarantees, unused overloads and constructors. Coverage against `main` (see [Tests](#tests)) shows what's reachable.

## Public API

Everything public is permanent: package validation flags removing or changing it, and users depend on it. Keep new public surface to what users need.

- Default to `internal` or `private`. Make a type or member public only when users call it, configure it, or implement or extend it, not because it's convenient for tests (tests go through the public API) or might be useful one day.
- Don't add public `const` fields for defaults or limits. Their values are compiled into callers, so changing them later doesn't reach users until they rebuild. Use a `private const` and state the value in the docs (for example "the first 50 rows"), or, if users need to read or change it, a property or method parameter.
- Prefer a method parameter or a setting on `ExcelImporterConfiguration` over a new public type for an option, and overloads over optional parameters (adding an optional parameter later is a breaking change).
- Name new APIs after the existing ones (`FindHeading`/`TryFindHeading`, `ValidateHeading`/`TryValidateHeading`, `Map(...).WithX(...)`), with a `Try` variant when failing is an expected outcome.
- Validate arguments and give actionable messages: say what was wrong and what to do instead.
- The extension points are `ICellConverter<T>`, `ICellTransformer`, `ICellFallback<T>` and `IColumnSelector`; maps are configured with `ExcelClassMap<T>.Map(...)` and the fluent methods. Add new behaviour as a fluent method or attribute on top of these, not by making a mapper, reader, factory, fallback, pipeline or map type public.
- Value methods are written once, for both `ValueMap<T>` and `ElementMap<T>`, never duplicated per map type. A method for any type is an instance method of `ValuePipeline<T, TSelf>` returning `TSelf` (like `WithTrim()`, `WithConverter(value => ...)` and `WithEmptyFallback(T?)`), so lambdas and `null` bind against the member's type. A method for particular types (`WithFormats` on dates), an overload for a nullable member's underlying type (`ICellConverter<int>` on an `int?` member) and the `ICellConverter<T>`/`ICellFallback<T>` overloads are extensions in `ValuePipelineExtensions` of the form `TMap M<TMap>(this ValuePipeline<X, TMap> map, ...) where TMap : ValuePipeline<X, TMap>`. The `ICellFallback<T>` overloads stay extensions so that `WithEmptyFallback(null)` binds to the value overload, which the compiler tries first.
- Until a major version ships, its breaking changes are allowed, but each one goes in the changelog, the migration guide and `CompatibilitySuppressions.xml` with its reason.

## Tests

Every change comes with tests, structured like the existing ones. There are two layers:

1. **Unit tests in `tests/ExcelMapper/`, one file per public type, mirroring `src/`.** A new public type `src/Foo/BarBaz.cs` gets `tests/ExcelMapper/Foo/BarBazTests.cs` with class `BarBazTests`. For example:
   - `src/ExcelColumnLetterAttribute.cs` → `tests/ExcelMapper/ExcelColumnLetterAttributeTests.cs`
   - `src/CellValue.cs` → `tests/ExcelMapper/CellValueTests.cs`
   - `src/ExcelRange.cs` → `tests/ExcelMapper/ExcelRangeTests.cs`

   New public members on an existing type are tested in that type's existing test file (for example a new `ValueMap<T>` method goes in `ValueMapTTests`, a new extension method in its extension class's tests).

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
- CSV test data is parsed like a real CSV file: the separator is auto-detected (`;` and `|` as well as `,`, so use `XlsxBuilder` for values containing them, or quote them), unquoted fields are trimmed, a row shorter than the heading has empty cells rather than missing columns, and a blank last line isn't a row. Calling `importer.ReadSheet()` again moves to the next sheet, so create a new importer for each read.
- Internal types (mappers, readers, fallbacks, bindings) are covered by mapping tests in `tests/Maps`, `tests/Fallbacks` and `tests/Transformers`. When a change removes or rewrites tests, compare coverage with `main` (`--collect:"XPlat Code Coverage" --settings coverage.runsettings`) and add mapping tests for anything no longer covered.

## Performance

Performance matters: the mapping path runs for every cell of every row. For any change to `src` that can affect it:

- Compare against `main`, preferably in-process (interleaved runs, best of several, `GC.GetAllocatedBytesForCurrentThread()` for allocations) or with the `Mapping` benchmark: `cd benchmarks && dotnet run -c Release -- --filter '*Mapping*' --job short`. BenchmarkDotNet timings on Apple Silicon can vary ~2x between runs; allocations are reliable, so compare those, and repeat timings.
- Report the numbers in the PR description.

Patterns to keep:

- Don't throw and catch exceptions for invalid cells: mappers use `TryParse` and return `CellConversion.Invalid`, and create the exception in `CreateException`, which is only called if the exception is used (the `MapperUtilities` helpers create it without throwing where they can).
- A lambda that captures locals allocates its closure when the locals' scope is entered, even if the lambda isn't used. Create such lambdas in a separate static method that takes the values as parameters.
- Use indexed `for` loops over `IList<T>` in per-cell code; `foreach` over an interface allocates an enumerator.
- Keep per-sheet state on the `ExcelSheet` (such as `GetCellReader`/`GetCellsReader` and `GetBoundMembers`), not in a dictionary on a map keyed by sheet: maps outlive sheets and are shared between importers and threads, so they must not hold sheets.

## Documentation

Update the docs in the same change as the code:

- `README.md` - overview, quick start and the documentation index. It is also the NuGet package page, so link to `docs/` pages with absolute GitHub URLs.
- `docs/*.md` - the topic pages (reading, mapping, types, errors, performance, API reference). Add new attributes and fluent methods to `docs/api.md` as well as the relevant topic page. Code samples use the public API and should compile; check what they claim (defaults, error messages) against the code or a test.
- `CHANGELOG.md` - add every user-visible change to the top (unreleased) section under Breaking changes, Added, Changed, Fixed or Performance, with a link to the PR. Check the PR's number once it's created (other PRs may have taken the one you expected), and keep Breaking changes ordered with the changes most users will hit first.
- Breaking changes (including behaviour changes, not only API changes) also go in the migration guide for that major version, for example `docs/migrating-to-4.0.md`: who is affected, before/now, and what to do, with before-and-after code where it helps.
- After a find-and-replace across the docs, read every changed line: a blanket replace once turned `UriKind.Absolute` into `UriKind.RelativeOrAbsolute`, including in the sentence giving the default.

## Releases and compatibility

- `PackageVersion` in `src/ExcelMapper.csproj` follows semantic versioning: a major version for breaking changes.
- Package validation (`EnablePackageValidation`) compares the package with `PackageValidationBaselineVersion` (the last release on NuGet) on every `dotnet pack`, including in CI, and fails on unintended API breaks. Intentional breaks are listed in `src/CompatibilitySuppressions.xml` with a comment explaining why; regenerate it with `dotnet pack src -c Release -p:GenerateCompatibilitySuppressionFile=true`. Regenerating drops the comment: copy it first, put it back, and add a line for the new break. After publishing a release, move the baseline to it and remove the suppressions it covered.
- Package validation only sees the API: behaviour changes still need the changelog and migration guide.

## Pull requests

- One feature or fix per PR, branched from `main`. Avoid stacking PRs on each other's branches: a PR merged into another feature branch doesn't reach `main`.
- CI (`.github/workflows/dotnetcore.yml`) runs for PRs into `main` only: build, tests with coverage, pack with package validation, the tests against the netstandard builds, and the trimmed test app.
- PRs are squash-merged.
- Don't force-push a branch that's been pushed: to bring in `main`, merge it into the branch.
- Before opening a PR, check:
  - `dotnet build src -c Release --no-incremental` has no warnings;
  - the tests pass against all three library builds (`net8.0`, `netstandard2.0`, `netstandard2.1`);
  - `dotnet pack src -c Release` passes package validation;
  - the trimmed app in `tests/TrimmingTests` publishes without trim warnings and passes, if mapping changed;
  - performance against `main` (see [Performance](#performance)), or a note in the PR saying why it can't change;
  - the changelog, the migration guide (for breaking changes), `docs/api.md` and the topic pages are updated.
- The PR description says what changed and why, how it was verified (test counts per build, trimmed app, pack), and the performance numbers.
