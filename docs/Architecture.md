# Architecture

We currently support
- One-to-one mapping: the value of the cell is transformed to a single value
- One-to-many mapping: the value of the cell is transformed to multiple values (e.g. a collection)

We do not yet support properly:
- Many-to-one mapping: the values of multiple cells are transformed to a single value
- Many-to-many mapping: the values of multiple cells are transformed to multiple values (e.g. a collection, dictionary, dynamic object)

## Pipeline
ExcelMapper follows a pipeline format. This depends on what is read.

### One-to-one
The class that handles this is `OneToOneMap`.
1. Read the value of a cell into a string using `ISingleCellValueReader.TryGetValue` for `OneToOneMap.CellReader`
- If the value is not found, bail out if the property is `Optional` or throw
- If the value is empty, it returns `IFallbackItem.PerformFallback` for `OneToOneMap.Pipeline.EmptyFallback` if it is present, or it continues
2. Perform any number of transformations on the string value using `ICellValueTransformer.TransformStringValue` for each value in `OneToOneMap.Transformers`
3. Map the string value into an object using `ICellValueMapper.MapCellValue` for each value in `OneToOneMap.Pipeline.CellValueMappers`. Each mapper is enumerated in order
- If the `Action` property of the returned `CellValueMapperResult` is `IgnoreResultAndContinueEnumeration`, then enumeration continues and ignores the success or error value of the result.
- If the `Action` property of the `CellValueMapperResult` is `UseResultAndContinueEnumeration`, then the enumeration continues and stores the success or error value of the result.
- If the `Action` property of the `CellValueMapperResult` is `UseResultAndStopEnumeration`, then the enumeration stops and stores the success or error value of the result.
- When the enumeration ends - either by reaching the end of `OneToOneMap.Pipeline.CellValueMappers` or by encountering `UseResultAndStopEnumeration` - then the last result used is handled.
If the mapping was successful, the `Value` property of the last `CellValueMapperResult` used is returned.
If the mapping was not successful, the `Exception` property of the last `CellValueMapperResult` used is passed to  `IFallbackItem.PerformFallback` using the property `OneToOneMap.Pipeline.InvalidFallback` if it is present, or an exception is thrown.

### One-to-many
1. Read the value of a cell into multiple strings using `IMultipleCellValuesReader.TryGetValues`