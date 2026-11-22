# CLAUDE.md

Guidance for working on ExcelMapper (NuGet package `ExcelDataReader.Mapping`), a library that maps Excel and CSV rows to .NET objects on top of ExcelDataReader.

## Layout

- `src/` - the library. Targets `netstandard2.0`, `netstandard2.1` and `net8.0`.
- `gen/` - the source generator for `ExcelMappingContext` (`[ExcelMappable]`), a netstandard2.0 Roslyn 4.8 generator packed into the library's package under `analyzers/dotnet/cs`. It must stay in step with auto-mapping: see [Auto-mapping and the source generator](#auto-mapping-and-the-source-generator). Tested in `tests/SourceGeneration`.
- `tests/` - xUnit tests (`net8.0`). See [Tests](#tests). `tests/TrimmingTests` and `tests/AotTests` are separate apps that CI publishes trimmed and with native AOT, and runs.
- `benchmarks/` - BenchmarkDotNet benchmarks, which can also run against the 3.x package. See [Performance](#performance) and `benchmarks/README.md`.
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

- The net8.0 build runs the trimming and native AOT analyzers (`IsTrimmable`, `IsAotCompatible`) and fails on their warnings. Annotate new reflection: put `[DynamicallyAccessedMembers]` on `Type` parameters and generic parameters where the caller knows the type (see `TrimmingConstants`), and `[RequiresUnreferencedCode]`/`[RequiresDynamicCode]` on code that discovers types at run time, carried up to the public API. Don't suppress warnings unless the justification is true. Changes to what mapping needs at run time should be checked with the trimmed app in `tests/TrimmingTests` and the native AOT app in `tests/AotTests` (see [docs/trimming.md](docs/trimming.md)). Native AOT only has the code of generic types it can see: create a generic type for a type known at compile time directly (`TypeRule<T>.Instance`, `new StructNumberMapper<int>()`), and one for a type found at run time with `GenericTypes.Create`, not `MakeGenericType`, with a matching `ExcelTypeRegistrar` method so an `ExcelMappingContext` can register it.

- Public types and members have XML doc comments. The documentation ships in the package (IntelliSense), and the build fails on missing or broken doc comments: undocumented public members, wrong `<param>` names, unresolved `cref`s.
- Validate arguments with `ThrowHelpers` (`ThrowIfNull`, `ThrowIfNegative`, ...) so parameter names match the API.
- Problems with a file or a mapping throw `ExcelMappingException`. Create it with the sheet, row and column where known so messages include the cell address (`at cell B11`).
- Nullable reference types are enabled.
- Attributes live in `src/` as `Excel<Name>Attribute.cs`; internal map interfaces in `src/Abstractions`, reading sheets in `src/Reading`, auto-mapping in `src/AutoMapping`, binding in `src/Binding`, mappers in `src/Mappers`, columns and readers in `src/Readers`, collection and dictionary builders in `src/Builders`, fallbacks in `src/Fallbacks`, helpers used across the library in `src/Utilities` (see the source layout in `docs/Architecture.md`). Everything in those folders is internal.
- Values are converted by `ValuePipeline<T>` as `T`, without boxing: mappers implement `ICellMapper<T>` (and `ICellMapper<T?>` for value types, so nullable members share them), fallbacks are `ICellFallback<T>`, and there's no untyped path. A new mapper implements `Convert` and `CreateException` (see `docs/Architecture.md`) and holds no state: settings that change how it parses, such as formats, go in the pipeline's `ConversionOptions` and reach it as `context.Options`.
- Errors collected by `ReadRowResults` travel as an `errors` list passed down through the maps, never as state on the sheet: code that can report a mapping error takes the list and adds to it when it isn't null.
- Walk a tree of maps with `MapTree` (`VisitCellMaps`, `FindOwner`), and add a map made of other maps by implementing `IParentMap`, rather than another switch over the kinds of map.
- A new way of choosing columns is a new kind of `Columns` (in its own file in `src/Readers`) that finds the columns and describes itself for the missing-column error and `ValidateHeading`, rather than a type check in `MapValidation` or `ExcelMappingException`.
- Auto-mapping decides what a type supports in its `TypeRule<T>`, `CollectionRule<TElement>` or `DictionaryRule<TKey, TValue>` (see `docs/Architecture.md`). Put a new type-specific rule there, generic over the type's arguments, rather than adding a `MakeGenericMethod` call elsewhere.
- Generic types created at run time with `MakeGenericType` need a `[DynamicDependency]` on the method that creates them, or the trimmer removes their constructors (check with the trimmed app).
- `Type.GetConstructor(types)` also matches constructors whose parameters the types are assignable to (asking for `ISet<T>` finds a constructor taking `IEnumerable<T>`). Use `CollectionUtilities.GetConstructorTaking` where the parameter type must match exactly.
- Delete code that can't be reached rather than keeping it "just in case": validation that every internal caller already guarantees, unused overloads and constructors. Coverage against `main` (see [Tests](#tests)) shows what's reachable.

## Auto-mapping and the source generator

Auto-mapping (reflection, with compiled delegates) and the source generator describe the same rules twice: auto-mapping decides at run time how each type is mapped, and the generator's `gen/TypeWalker.cs` decides at compile time which types an `ExcelMappingContext` registers so that the same decisions work with native AOT. **Change them together, in the same PR, in both directions.** A rule changed in only one of them breaks native AOT apps (a type that isn't registered throws `NotSupportedException`) without failing anything under the JIT.

They mirror each other here (each place has a comment pointing at the other):

| Auto-mapping | `TypeWalker` |
|---|---|
| `AutoMapper.TryCreateMap`: a value, then a dictionary, a collection, an object | `Visit`, in the same order |
| `TypeRule<T>.FindDefaultMappers`: which types have a built-in conversion (built-in types, enums, numbers, `IConvertible`, `IParsable<T>`) | `s_builtInTypes`, `HasConversion`, `CanConstructObject` |
| `DictionaryRule.TryGetKeyValueTypes` | `TryGetKeyValueTypes` |
| `ReflectionUtilities.GetElementTypeOrEnumerableType` | `TryGetElementType`, `GetSingleAdd` |
| `AutoMapper.TryAddMembers`, `ShouldAutoMapProperty`, `[ExcelIgnore]` | `GetMappedMembers` |
| each `GenericTypes.Create` call (a generic type made for a type found at run time) | an `ExcelTypeRegistrar` method that registers it, which `TypeWalker` calls |
| `CompiledAccessors.GetSetter`: which members are set, and how (a property's public setter, a field that isn't readonly) | `AddSetter`, which registers a generated setter for each |
| `ConstructionPlan<T>.Create` and `ConstructorUtilities`: which constructors create a type | `AddConstructors`, which registers a generated call for each public constructor with parameters |

So:

- **When auto-mapping changes** how a kind of type is mapped, which members it reads or sets, which constructors it calls, or adds a `GenericTypes.Create` call: update `TypeWalker` (adding an `ExcelTypeRegistrar` method for a new generic type), and add cases to `tests/SourceGeneration/ExcelMappableGeneratorTests.cs` (the registrations written) and `GeneratedContextTests.cs` (a generated context maps rows as `ReadRows<T>()` does).
- **When the generator changes** which types it registers or how it walks members: check it still matches auto-mapping as above, rather than making the generator decide differently.
- A generic type made at run time must go through `GenericTypes.Create`, never `MakeGenericType` directly, so that a context can register it. Types known at compile time are created directly (see [Code conventions](#code-conventions)).
- If the change matters with native AOT (a new kind of type, member or collection), add a case to the generated context in `tests/AotTests`, which fails on any trimming or AOT warning and on a missing registration.
- The netstandard builds map some types differently (for example, they don't parse `IParsable<T>` types): `TypeWalker` checks what the library it compiles against supports, and `ExcelMappableGeneratorTests` has expectations for both.

## Public API

Everything public is permanent: package validation flags removing or changing it, and users depend on it. Keep new public surface to what users need.

The public API of each target framework is listed in `src/PublicAPI/<framework>/` (Microsoft.CodeAnalysis.PublicApiAnalyzers), and the build fails if the API and the lists differ, so every change to the public API shows in the diff. New APIs go in `PublicAPI.Unshipped.txt`; at a release, move its lines to `PublicAPI.Shipped.txt`. To update the lists, apply the analyzer's "Add to public API" fix in an IDE, or for each framework run `dotnet format analyzers src/ExcelMapper.csproj --diagnostics RS0016 --severity info` with `<TargetFrameworks>` temporarily set to that framework alone (`dotnet format` only fixes the first). For a removed API (RS0017), delete its line. Warnings are errors in `src`, so fix a warning rather than suppressing it.

- Default to `internal` or `private`. Make a type or member public only when users call it, configure it, or implement or extend it, not because it's convenient for tests (tests go through the public API) or might be useful one day.
- Don't add public `const` fields for defaults or limits. Their values are compiled into callers, so changing them later doesn't reach users until they rebuild. Use a `private const` and state the value in the docs (for example "the first 50 rows"), or, if users need to read or change it, a property or method parameter.
- Prefer a method parameter or a setting on `ExcelImporterConfiguration` over a new public type for an option, and overloads over optional parameters (adding an optional parameter later is a breaking change, and the analyzer rejects two overloads of a method that both have optional parameters).
- Name new APIs after the existing ones (`FindHeading`/`TryFindHeading`, `ValidateHeading`/`TryValidateHeading`, `Map(...).WithX(...)`), with a `Try` variant when failing is an expected outcome.
- Validate arguments and give actionable messages: say what was wrong and what to do instead.
- The extension points are `ICellConverter<T>`, `ICellTransformer`, `ICellFallback<T>` and `IColumnSelector`; maps are configured with `ExcelClassMap<T>.Map(...)` and the fluent methods. Add new behaviour as a fluent method or attribute on top of these, not by making a mapper, reader, builder, fallback, pipeline or map type public.
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

- Compare against `main`, preferably in-process (interleaved runs, best of several, `GC.GetAllocatedBytesForCurrentThread()` for allocations) or with the `Mapping` benchmark: `cd benchmarks && dotnet run -c Release -- --filter '*.Mapping.*' --job short` (see `benchmarks/README.md`). BenchmarkDotNet timings on Apple Silicon can vary ~2x between runs; allocations are reliable, so compare those, and repeat timings.
- Report the numbers in the PR description.

Patterns to keep:

- Don't throw and catch exceptions for invalid cells: mappers use `TryParse` and return `CellConversion.Invalid`, and create the exception in `CreateException`, which is only called if the exception is used (the `MapperUtilities` helpers create it without throwing where they can).
- A lambda that captures locals allocates its closure when the locals' scope is entered, even if the lambda isn't used. Create such lambdas in a separate static method that takes the values as parameters.
- Use indexed `for` loops over `IList<T>` in per-cell code; `foreach` over an interface allocates an enumerator.
- Maps are frozen the first time they're used with a sheet (`IMap.Freeze`), so nothing worked out from them needs to be rebuilt. Anything new that configures a map must throw once it's frozen (`ThrowIfFrozen`, or a `NonNullCollection<T>`), and a new kind of map must freeze the maps and pipelines under it.
- Keep per-sheet state in the sheet's `SheetBindings` (through `SheetView.FindColumn`, `GetCellsReader` and `GetBoundMembers`), not in a dictionary on a map keyed by sheet: maps outlive sheets and are shared between importers and threads, so they must not hold sheets.

## Documentation

Update the docs in the same change as the code:

- `README.md` - overview, quick start and the documentation index. It is also the NuGet package page, so link to `docs/` pages with absolute GitHub URLs.
- `docs/*.md` - the topic pages (reading, mapping, types, errors, performance, API reference). Add new attributes and fluent methods to `docs/api.md` as well as the relevant topic page. Code samples use the public API and must compile: `tests/Docs/DocSamplesTests.cs` compiles every ```` ```csharp ```` block in `README.md` and `docs/` (except `Architecture.md`) as a C# script, after a prelude that declares `importer`, `sheet`, `Employee` and the other names samples use without declaring. Show class-map lines as `importer.Configuration.RegisterClassMap<T>(c => c.Map(...))` or inside a class map, not as a bare `Map(...)`; in before-and-after samples, start the old code with a `// 3.x` comment and the new with `// 4.0` (only the 4.0 code is compiled); fence a block that isn't meant to compile as ```` ```csharp no-compile ````. Check what samples claim (defaults, error messages) against the code or a test, as they're compiled but not run.
- `CHANGELOG.md` - add every user-visible change to the top (unreleased) section under Breaking changes, Added, Changed, Fixed or Performance, with a link to the PR. Check the PR's number once it's created (other PRs may have taken the one you expected), and keep Breaking changes ordered with the changes most users will hit first.
- Breaking changes (including behaviour changes, not only API changes) also go in the migration guide for that major version, for example `docs/migrating-to-4.0.md`: who is affected, before/now, and what to do, with before-and-after code where it helps.
- After a find-and-replace across the docs, read every changed line: a blanket replace once turned `UriKind.Absolute` into `UriKind.RelativeOrAbsolute`, including in the sentence giving the default.

## Releases and compatibility

- `PackageVersion` in `src/ExcelMapper.csproj` follows semantic versioning: a major version for breaking changes.
- Package validation (`EnablePackageValidation`) compares the package with `PackageValidationBaselineVersion` (the last release on NuGet) on every `dotnet pack`, including in CI, and fails on unintended API breaks. Intentional breaks are listed in `src/CompatibilitySuppressions.xml` with a comment explaining why; regenerate it with `dotnet pack src -c Release -p:GenerateCompatibilitySuppressionFile=true`. Regenerating drops the comment: copy it first, put it back, and add a line for the new break. After publishing a release, move the baseline to it and remove the suppressions it covered.
- Package validation only sees the API: behaviour changes still need the changelog and migration guide.

To release a version:

1. In a pull request: set `PackageVersion` in `src/ExcelMapper.csproj`, give the version's changelog heading its date, and move the lines of each `src/PublicAPI/<framework>/PublicAPI.Unshipped.txt` (except `#nullable enable`) to the `PublicAPI.Shipped.txt` beside it.
2. After it's merged, tag `main` with the version and push the tag: `git tag v4.0.0 && git push origin v4.0.0`. `.github/workflows/release.yml` checks the tag matches `PackageVersion`, builds, runs the same tests as CI and the trimmed app, packs (with package validation), and publishes the package and its symbols to NuGet with Trusted Publishing (a policy on the nuget.org account for this repository and `release.yml`, and the `NUGET_USER` repository secret).
3. Then, in a pull request, set `PackageValidationBaselineVersion` to the version just released and remove the suppressions in `src/CompatibilitySuppressions.xml` that it covered.

## Pull requests

- One feature or fix per PR, branched from `main`. Avoid stacking PRs on each other's branches: a PR merged into another feature branch doesn't reach `main`.
- CI (`.github/workflows/dotnetcore.yml`) runs for PRs into `main` only: build, tests with coverage, pack with package validation, the tests against the netstandard builds, the trimmed test app, and the native AOT test app.
- PRs are squash-merged.
- Don't force-push a branch that's been pushed: to bring in `main`, merge it into the branch.
- Before opening a PR, check:
  - `dotnet build src -c Release --no-incremental` has no warnings;
  - the tests pass against all three library builds (`net8.0`, `netstandard2.0`, `netstandard2.1`);
  - `dotnet pack src -c Release` passes package validation;
  - the trimmed app in `tests/TrimmingTests` publishes without trim warnings and passes, and the native AOT app in `tests/AotTests` passes (`dotnet publish tests/AotTests -c Release -r <rid> -o out && out/AotTests`), if mapping changed;
  - auto-mapping and the source generator's `TypeWalker` still agree, if either changed (see [Auto-mapping and the source generator](#auto-mapping-and-the-source-generator));
  - performance against `main` (see [Performance](#performance)), or a note in the PR saying why it can't change;
  - the changelog, the migration guide (for breaking changes), `docs/api.md` and the topic pages are updated.
- The PR description says what changed and why, how it was verified (test counts per build, trimmed app, pack), and the performance numbers.
