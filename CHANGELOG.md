# Changelog

All notable changes to ExcelMapper (NuGet package [`ExcelDataReader.Mapping`](https://www.nuget.org/packages/ExcelDataReader.Mapping/)) are documented here. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project follows [Semantic Versioning](https://semver.org/).

## [4.0.0] - Unreleased

4.0 contains breaking changes. See [Migrating to 4.0](docs/migrating-to-4.0.md) for what to check when upgrading from 3.x.

### Breaking changes

- `object` and `IConvertible` members, and `object` values in dictionaries, `ExpandoObject` and non-generic collections, receive the value as Excel stores it (`double`, `DateTime`, `bool`, `string`, ...) instead of its text. ([#124](https://github.com/hughbe/excel-mapper/pull/124))
- A fractional number (such as `2.7`) mapped to an integer type is invalid instead of being truncated (.NET 8) or rounded (netstandard), matching text cells. ([#127](https://github.com/hughbe/excel-mapper/pull/127))
- `ReadRow` and `TryReadRow` read the heading first if it hasn't been read, like `ReadRows`, instead of mapping the heading row as data. ([#119](https://github.com/hughbe/excel-mapper/pull/119))
- Sheet names are matched case-insensitively by `ReadSheet(name)` and `TryReadSheet(name, ...)`, as in Excel. ([#116](https://github.com/hughbe/excel-mapper/pull/116))
- Error messages give Excel cell addresses and one-based rows (`at cell B11`, `on row 11`) instead of zero-based row indexes. ([#117](https://github.com/hughbe/excel-mapper/pull/117))
- `ExcelRange` is a `readonly struct`; `default(ExcelRange)` (and the value `TryParse` returns on failure) means every row and column instead of an empty range; `Parse` rejects non-ASCII letters, signs, and rows or columns beyond Excel's limits. ([#117](https://github.com/hughbe/excel-mapper/pull/117))
- With a `DataRange` that doesn't start at the first column, `ExcelHeading.GetColumnIndex` and `GetColumnName` use positions in the sheet rather than in the range. ([#116](https://github.com/hughbe/excel-mapper/pull/116))
- Splitting a cell with a trailing separator (`"a,b,"`) includes the trailing empty value, as `string.Split` does, for arrays and lists alike; an empty cell gives no values. ([#116](https://github.com/hughbe/excel-mapper/pull/116))
- `RegisterClassMap` replaces a class map that auto-mapping registered implicitly instead of throwing. ([#120](https://github.com/hughbe/excel-mapper/pull/120))
- The netstandard builds no longer expose a public `CallerArgumentExpressionAttribute` type, which was a polyfill accidentally made public (and ignored by the compiler).

### Added

- Records and types without a parameterless constructor are created through a constructor whose parameters match mapped members; read-only properties can be mapped through constructor parameters. ([#118](https://github.com/hughbe/excel-mapper/pull/118))
- `[ExcelRowNumber]` and `MapRowNumber(...)` set a member to the row's number, as Excel shows it; `[ExcelExtraColumns]` and `MapExtraColumns(...)` fill a dictionary with the columns no other member reads. ([#136](https://github.com/hughbe/excel-mapper/pull/136))
- `ExcelSheet.ReadRowResults<T>()` reports every invalid cell in every row instead of stopping at the first. ([#126](https://github.com/hughbe/excel-mapper/pull/126))
- `ExcelSheet.ValidateHeading<T>()` and `TryValidateHeading<T>()` report every missing column before reading rows. ([#128](https://github.com/hughbe/excel-mapper/pull/128))
- `ExcelImporter.ReadRows<T>(path or stream)` reads the first sheet in one call. ([#129](https://github.com/hughbe/excel-mapper/pull/129))
- `[ExcelColumnLetter("C")]` and `.WithColumnLetter("C")` map columns by Excel letter. ([#129](https://github.com/hughbe/excel-mapper/pull/129))
- `ExcelImporterConfiguration.Culture` sets the culture used for parsing and converting values to text. ([#127](https://github.com/hughbe/excel-mapper/pull/127))
- `ExcelImporterConfiguration.IgnoreSeparatorsInColumnNames` matches `FirstName` to columns such as "First Name"; errors for a missing column suggest a similarly named one. ([#122](https://github.com/hughbe/excel-mapper/pull/122))
- `ExcelImporter` constructors that take an `ExcelReaderConfiguration` (passwords, CSV separators, fallback encodings). ([#122](https://github.com/hughbe/excel-mapper/pull/122))
- `[ExcelOffset("+00:00")]`, `.WithOffset(TimeSpan)` and `DateTimeOffsetMapper.Offset` choose the offset of `DateTimeOffset` members. ([#131](https://github.com/hughbe/excel-mapper/pull/131))
- `ExcelMappingException.CellAddress`, `ColumnName` and `Member`. ([#117](https://github.com/hughbe/excel-mapper/pull/117), [#130](https://github.com/hughbe/excel-mapper/pull/130))
- `ExcelRange` equality operators, `ToString()` (an Excel address such as `"A1:B10"`) and `$` absolute references. ([#117](https://github.com/hughbe/excel-mapper/pull/117))
- `WithColumnsMatching(Func<string, bool>)`, `WithColumnsMatching(Regex)` and `WithColumnNameMatching(Regex)`. ([#119](https://github.com/hughbe/excel-mapper/pull/119))
- `CellMapperResult.Invalid(Func<Exception>)` for mappers that create exceptions only when needed. ([#121](https://github.com/hughbe/excel-mapper/pull/121))
- Numbers in cells without a date format map to `DateTime`, `DateTimeOffset` and `DateOnly` (Excel serial dates), `TimeOnly` (fraction of a day) and `TimeSpan` (days); time cells map to `TimeSpan`; `bool` accepts "yes" and "no". ([#123](https://github.com/hughbe/excel-mapper/pull/123))
- Sheets without a heading can be auto-mapped when columns aren't looked up by name, for example primitives or `[ExcelColumnIndex]` members. ([#119](https://github.com/hughbe/excel-mapper/pull/119))
- `ExcelMappingException` is thrown, instead of `NullReferenceException`, when an empty value can't be assigned to a non-nullable value type. ([#116](https://github.com/hughbe/excel-mapper/pull/116))

### Changed

- Code page encodings are registered automatically; `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)` is no longer needed. The netstandard builds now depend on `System.Text.Encoding.CodePages`. ([#122](https://github.com/hughbe/excel-mapper/pull/122))
- Trimming only changes text, so cells Excel stores as dates or numbers keep their type. ([#122](https://github.com/hughbe/excel-mapper/pull/122))
- Exceptions thrown by a constructor are wrapped in `ExcelMappingException`. ([#118](https://github.com/hughbe/excel-mapper/pull/118))
- Building an `ExcelMappingException` message no longer reads the sheet's heading. ([#130](https://github.com/hughbe/excel-mapper/pull/130))
- Numbers are parsed the same way on every build. On the netstandard builds, `[ExcelNumberStyle]` now takes effect (previously it was ignored there), and text is parsed with `NumberStyles.Number`, so thousands separators such as `"1,234"` are accepted for integers as they were on .NET 8. Floating-point types also accept exponents such as `"1E5"` on every build (previously only on netstandard). ([#135](https://github.com/hughbe/excel-mapper/pull/135))
- Updated ExcelDataReader to 3.9.0 and ExcelNumberFormat to 1.2.0. ([#115](https://github.com/hughbe/excel-mapper/pull/115))
- The documentation moved from the README into [`docs/`](docs). ([#132](https://github.com/hughbe/excel-mapper/pull/132))

### Fixed

- A `DataRange` that skips leading columns read values from the wrong columns. ([#116](https://github.com/hughbe/excel-mapper/pull/116))
- Column matching threw when the `DataRange` didn't include every column; `GetFirstColumnMatchingIndex` returned the alphabetically first match. ([#116](https://github.com/hughbe/excel-mapper/pull/116))
- The `DataRange` row start was ignored for sheets without a heading; `SkipBlankLines` checked columns outside the `DataRange`. ([#116](https://github.com/hughbe/excel-mapper/pull/116))
- `SkipBlankLines` read past the end of the `DataRange`; `ReadRows(startIndex, count)` returned later rows when the reader was already past `startIndex`; CSV sheets reported `VeryHidden`. ([#113](https://github.com/hughbe/excel-mapper/pull/113))
- Sharing a class map between importers on different threads could corrupt results, and kept every sheet it had read in memory. ([#120](https://github.com/hughbe/excel-mapper/pull/120))
- Changing a map's `ReaderFactory` after reading kept using the old column. ([#120](https://github.com/hughbe/excel-mapper/pull/120))
- Building an error message could read the heading as a side effect. ([#130](https://github.com/hughbe/excel-mapper/pull/130))
- On the netstandard builds, argument exceptions had no parameter name, and `null` arguments to string validation threw `ArgumentException` instead of `ArgumentNullException`.

### Performance

- Invalid values no longer throw and catch an exception for every cell: 7x faster with a fallback. ([#121](https://github.com/hughbe/excel-mapper/pull/121))
- Compiled setters and constructors, and fewer per-row and per-cell allocations. ([#121](https://github.com/hughbe/excel-mapper/pull/121), [#127](https://github.com/hughbe/excel-mapper/pull/127))
- `ReadRowResults` costs the same as `ReadRows` for valid data. ([#126](https://github.com/hughbe/excel-mapper/pull/126))

### Platform notes

The package contains `net8.0`, `netstandard2.1` and `netstandard2.0` builds. Some features need the `net8.0` build (used by .NET 8 and later):

- Mapping `DateOnly`, `TimeOnly`, `Half`, `Int128`, `UInt128`, `nint`, `nuint`, `BigInteger`, `Complex` and other `IParsable<T>` / `INumberBase<T>` types.
- `.WithNumberStyle(...)` and `.WithFormatProvider(...)` for numeric members (use `[ExcelNumberStyle]` and `ExcelImporterConfiguration.Culture` on the netstandard builds).
- `.WithFormats(...)` and related methods for `DateOnly` and `TimeOnly`, and mapping `IReadOnlySet<T>` members.

The test suite now also runs against the netstandard builds in CI.

## Earlier versions

Changes before 4.0.0 are listed in the [GitHub history](https://github.com/hughbe/excel-mapper/commits/main) and on [NuGet](https://www.nuget.org/packages/ExcelDataReader.Mapping/#versions-body-tab).
