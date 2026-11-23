# Benchmarks

[BenchmarkDotNet](https://benchmarkdotnet.org) benchmarks for ExcelMapper. The `Mapping` benchmarks measure common scenarios on 50,000 rows of in-memory CSV (or xlsx for `TypedNumbers` and `TypedDates`), so that reading the file costs the same in every benchmark:

| Benchmark | What it maps |
|---|---|
| `ReaderOnly` | Nothing: reads every cell with ExcelDataReader directly. The floor for any mapping, and the baseline for the ratios. |
| `AutoMappedClass` | An auto-mapped 10-member class: strings, numbers, a date, a boolean, a `Guid`, an enum and a nullable number. |
| `FluentClassMap` | A class map with `WithTrim`, fallbacks, `WithConverter(value => ...)`, `WithMapping` and a column name. |
| `InvalidValuesWithFallback` | A number column where half the values are invalid and get an `[ExcelInvalidValue]`. |
| `TypedNumbers` | Number cells from an xlsx file, which ExcelDataReader returns as `double`. |
| `SplitCollection` | A `List<int>` split from each cell, such as `"1,2,3"`. |
| `DictionaryOfColumns` | Each row as a `Dictionary<string, int>` of its columns. |
| `TypedDates` | Date cells from an xlsx file, which ExcelDataReader returns as `DateTime`. |
| `SkipBlankLines` | Rows with `SkipBlankLines` set, where every other row is blank. |
| `OptionalMissingColumn` | `AutoMappedClass` with an optional member whose column the sheet doesn't have. |
| `AutoMapping` | One row of a new type with collection, nested, enum and nullable members: the cost of auto-mapping a type, which each importer does the first time it reads it. |
| `AutoMappedClassWithCulture` | `AutoMappedClass` with `ExcelImporterConfiguration.Culture` set (4.0). |
| `AutoMappedStruct` | `AutoMappedClass` with each row a struct. |
| `AutoMappedRecord` | `AutoMappedClass` as a positional record (4.0). |
| `RowResultsClass`, `RowResultsInvalidValues` | `ReadRowResults` on valid rows, and on rows half of which are invalid, without a fallback (4.0). |
| `TypedConverter` | An `ICellConverter<int>` on the invalid values (4.0). |
| `Region` | `AutoMappedClass` read as a region of the sheet (4.0). |

The `NativeAot` benchmarks run each benchmark with the JIT and compiled as a native AOT app, on the current .NET, mapping the same 50,000 rows with a class map from a generated `ExcelMappingContext` (`BenchmarkContext`), as native AOT apps do:

| Benchmark | What it maps |
|---|---|
| `ReaderOnly` | Nothing: reads every cell with ExcelDataReader directly. |
| `Class`, `Record`, `Struct` | The 10-member class of `AutoMappedClass`, as a class, a positional record and a struct. |

## Running them

```bash
cd benchmarks
dotnet run -c Release -- --filter '*.Mapping.*'
```

Add `--job short` for a quicker, less precise run (except for `NativeAot`, which sets its own jobs: use `--iterationCount 5 --warmupCount 2`). The `NativeAot` benchmarks need what native AOT needs to compile (on Linux, `clang` and `zlib1g-dev`), and take a minute or two to build:

```bash
dotnet run -c Release -- --filter '*.NativeAot.*'
```

On Apple Silicon, timings can vary by up to 2x between runs: compare allocations, which are reliable, and repeat timings.

## Comparing with 3.x

Set `EXCELMAPPER_VERSION` to run the `Mapping` benchmarks against that version of the [`ExcelDataReader.Mapping`](https://www.nuget.org/packages/ExcelDataReader.Mapping/) package instead of `src/`. The benchmarks for 4.0 features are left out, and ExcelDataReader is the version `src/` uses, so the difference is ExcelMapper's alone:

```bash
EXCELMAPPER_VERSION=3.0.3 dotnet run -c Release -- --filter '*.Mapping.*'
```

CI builds this too, so it keeps compiling.

## 4.0 compared with 3.0.3

Apple M4, .NET 9.0.10, BenchmarkDotNet 0.15.8, default job, 4.0 at [#218](https://github.com/hughbe/excel-mapper/pull/218). 3.0.3 was run twice, and each row gives its faster run. Allocations include ExcelDataReader's 165.6 MB for reading the file (`ReaderOnly`), which is most of the time too.

| Benchmark | 3.0.3 | 4.0 | Faster | Allocated (3.0.3 → 4.0) |
|---|---:|---:|---:|---:|
| `ReaderOnly` | 94.9 ms | 93.8 ms | | 165.6 MB → 165.6 MB |
| `AutoMappedClass` | 137.6 ms | 112.3 ms | 1.2x | 243.5 MB → 170.6 MB |
| `AutoMappedStruct` | 137.0 ms | 112.5 ms | 1.2x | 243.5 MB → 165.6 MB |
| `FluentClassMap` | 118.4 ms | 103.5 ms | 1.1x | 211.8 MB → 168.3 MB |
| `InvalidValuesWithFallback` | 216.6 ms | 24.2 ms | 9.0x | 85.7 MB → 59.7 MB |
| `TypedNumbers` | 61.2 ms | 58.3 ms | 1.05x | 60.9 MB → 36.7 MB |
| `SplitCollection` | 44.7 ms | 37.5 ms | 1.2x | 94.5 MB → 74.9 MB |
| `DictionaryOfColumns` | 58.7 ms | 45.1 ms | 1.3x | 124.9 MB → 99.5 MB |
| `AutoMapping` | 44.5 µs | 11.6 µs | 3.8x | 66.2 KB → 41.9 KB |
| `TypedDates` | 53.2 ms | 52.3 ms | 1.02x | 46.5 MB → 29.7 MB |
| `SkipBlankLines` | 38.2 ms | 34.5 ms | 1.1x | 121.4 MB → 107.0 MB |
| `OptionalMissingColumn` | 101.2 ms | 97.4 ms | 1.04x | 185.1 MB → 167.5 MB |

Most of each mapping benchmark's time is ExcelDataReader reading the file (`ReaderOnly`), which is the same for both versions. Beyond it, mapping the 10-member class takes 18.5 ms instead of 42.7 ms (2.3 times faster) and allocates 5.1 MB instead of 77.9 MB (93% less); a struct allocates nothing beyond reading the file. 3.0.3 threw and caught an exception for each of the 25,000 invalid values; 4.0 throws none.

An earlier version of this table gave 3.0.3's `AutoMappedClass` as 494.8 ms, and 4.0 as 4.3 times faster: two later runs of 3.0.3 took 138 to 142 ms, so that run was wrong.

## Native AOT compared with the JIT

Apple M4, .NET 9.0.10, BenchmarkDotNet 0.15.8, default job, with generated setters and constructors ([#202](https://github.com/hughbe/excel-mapper/pull/202)). Each benchmark creates the importer, so its time includes ExcelDataReader opening the file.

| Benchmark | JIT | Native AOT | Ratio | Allocated (JIT → native AOT) |
|---|---:|---:|---:|---:|
| `ReaderOnly` | 95.2 ms | 173.7 ms | 1.82 | 161.67 MB → 161.67 MB |
| `Class` | 112.3 ms | 198.2 ms | 1.77 | 166.64 MB → 166.67 MB |
| `Record` | 114.1 ms | 198.7 ms | 1.74 | 166.64 MB → 166.64 MB |
| `Struct` | 112.0 ms | 197.0 ms | 1.76 | 161.68 MB → 161.71 MB |

Native AOT allocates the same as the JIT. Nearly all of its extra time is ExcelDataReader reading the file (`ReaderOnly`), which native AOT runs about 80% slower, as it can't optimize from how the code runs as the JIT does. Mapping itself adds 24.5 ms with native AOT and 17.1 ms with the JIT. Before generated setters and constructors, native AOT set members by reflection: a struct allocated 237 MB instead of 162 MB, a class 193 MB instead of 167 MB.
