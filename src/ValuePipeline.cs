using System.Reflection;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Transformers;

namespace ExcelMapper;

/// <summary>
/// Reads a single cell of an excel sheet and maps the value of the cell to the
/// type of the property or field.
/// </summary>
public class ValuePipeline
{
    // Pipelines are typed: see ValuePipeline<T>.
    private protected ValuePipeline()
    {
    }

    /// <summary>
    /// Gets the list of objects that take the initial string value read from a cell and
    /// modifies the string value. This is useful for things like trimming the string value
    /// before mapping it.
    /// </summary>
    internal IList<Abstractions.ICellTransformer> Transformers { get; } = new NonNullCollection<Abstractions.ICellTransformer>();

    /// <summary>
    /// Gets the pipeline of items that take the initial string value read from a cell and
    /// converts the string value into the type of the property or field. The items form
    /// a pipeline: if a mapper fails to parse or map the cell value, the next item is used.
    /// </summary>
    internal IList<ICellMapper> Mappers { get; } = new NonNullCollection<ICellMapper>();

    /// <summary>
    /// Gets or sets an object that handles mapping a cell value to a property or field if the value of the
    /// cell is empty. For example, you can provide a fixed value to return if the value of the cell
    /// is empty.
    /// </summary>
    internal IFallbackItem? EmptyFallback
    {
        get => _emptyFallback;
        set
        {
            _emptyFallback = value;
            _fallbackVersion++;
        }
    }

    private IFallbackItem? _emptyFallback;

    /// <summary>
    /// Gets or sets an object that handles mapping a cell value to a property or field if all items
    /// in the mapper pipeline failed to map the value to the property or field. For example, you can
    /// provide a fixed value to return if the value of the cell is invalid.
    /// </summary>
    internal IFallbackItem? InvalidFallback
    {
        get => _invalidFallback;
        set
        {
            _invalidFallback = value;
            _fallbackVersion++;
        }
    }

    private IFallbackItem? _invalidFallback;

    private int _fallbackVersion;

    /// <summary>
    /// Changes whenever the transformers, mappers or fallbacks change, so that plans built from the
    /// pipeline (see <see cref="OneToOneMap{T}"/>) are rebuilt.
    /// </summary>
    internal int Version
        => ((NonNullCollection<Abstractions.ICellTransformer>)Transformers).Version + ((NonNullCollection<ICellMapper>)Mappers).Version + _fallbackVersion;

    /// <summary>
    /// Processes a cell value through the complete transformation and mapping pipeline to produce a final property value.
    /// This method orchestrates the entire value mapping process: transforming the raw cell value, running it through
    /// the mapper pipeline, and applying fallback strategies when needed.
    /// </summary>
    /// <param name="pipeline">The value pipeline containing transformers, mappers, and fallback items to use.</param>
    /// <param name="sheet">The Excel sheet being read, used for context in transformers and error messages.</param>
    /// <param name="rowIndex">The zero-based index of the row being processed, used for error reporting.</param>
    /// <param name="readResult">The result of reading a cell, containing the raw value and column information.</param>
    /// <param name="preserveFormatting">Whether to preserve Excel formatting when reading string values.</param>
    /// <param name="member">The property or field being mapped to, used for error messages. Can be null.</param>
    /// <returns>
    /// The final mapped value after applying transformers, mappers, and fallbacks. 
    /// May return null if empty fallback returns null, or throw an exception if fallbacks are configured to throw.
    /// </returns>
    /// <remarks>
    /// <para>Processing order:</para>
    /// <list type="number">
    /// <item><description>Apply all cell value transformers (e.g., trimming)</description></item>
    /// <item><description>Check if value is empty and apply empty fallback if configured</description></item>
    /// <item><description>Run value through mapper pipeline until one succeeds or all fail</description></item>
    /// <item><description>If all mappers fail, apply invalid fallback if configured</description></item>
    /// <item><description>Return the final mapped value or throw if no fallback handled the failure</description></item>
    /// </list>
    /// </remarks>
    internal static object? GetPropertyValue(
        ValuePipeline pipeline,
        ExcelSheet sheet,
        int rowIndex,
        ReadCellResult readResult,
        bool preserveFormatting,
        MemberInfo? member)
    {
        readResult = ApplyTransformers(pipeline, sheet, rowIndex, readResult, preserveFormatting, member);
        if (readResult.IsEmpty() && pipeline.EmptyFallback != null)
        {
            return pipeline.EmptyFallback.PerformFallback(sheet, rowIndex, readResult, null, member);
        }

        CellMapperResult? finalResult = null;
        var mappers = pipeline.Mappers;
        for (int i = 0; i < mappers.Count; i++)
        {
            var mapper = mappers[i];
            var result = mapper is IContextCellMapper contextMapper
                ? contextMapper.Map(readResult, sheet, rowIndex, member)
                : mapper.Map(readResult);
            if (result.Action != CellMapperResult.HandleAction.IgnoreResultAndContinueMapping)
            {
                finalResult = result;
            }

            if (result.Action == CellMapperResult.HandleAction.UseResultAndStopMapping)
            {
                break;
            }
        }


        if ((finalResult == null || !finalResult.Value.Succeeded) && pipeline.InvalidFallback != null)
        {
            var exception = FallbackUsesException(pipeline.InvalidFallback) ? finalResult?.Exception : null;
            return pipeline.InvalidFallback.PerformFallback(sheet, rowIndex, readResult, exception, member);
        }

        return finalResult?.Value;
    }

    /// <summary>
    /// Applies the pipeline's transformers to a cell.
    /// </summary>
    internal static ReadCellResult ApplyTransformers(ValuePipeline pipeline, ExcelSheet sheet, int rowIndex, ReadCellResult readResult, bool preserveFormatting, MemberInfo? member)
    {
        // Use indexed loops: enumerating an IList through its interface allocates for every cell.
        var transformers = pipeline.Transformers;
        for (int i = 0; i < transformers.Count; i++)
        {
            var transformer = transformers[i];

            // Trimming only changes text. Skip it for other values, such as dates and numbers,
            // so that they keep their type rather than being converted to a string and parsed again.
            if (transformer is TrimStringCellTransformer && !preserveFormatting && readResult.GetValue() is not (string or null))
            {
                continue;
            }

            readResult = transformer is CellTransformerAdapter cellTransformer
                ? cellTransformer.Transform(readResult, sheet, rowIndex, member)
                : new ReadCellResult(readResult.ColumnIndex, transformer.TransformStringValue(sheet, rowIndex, readResult), preserveFormatting);
        }

        return readResult;
    }

    /// <summary>
    /// Applies transformers from a plan (see <see cref="OneToOneMap{T}"/>) to a cell.
    /// </summary>
    internal static ReadCellResult ApplyTransformers(Abstractions.ICellTransformer[] transformers, ExcelSheet sheet, int rowIndex, ReadCellResult readResult, bool preserveFormatting, MemberInfo? member)
    {
        for (int i = 0; i < transformers.Length; i++)
        {
            var transformer = transformers[i];

            // Trimming only changes text: see ApplyTransformers(ValuePipeline, ...).
            if (transformer is TrimStringCellTransformer && !preserveFormatting && readResult.GetValue() is not (string or null))
            {
                continue;
            }

            readResult = transformer is CellTransformerAdapter cellTransformer
                ? cellTransformer.Transform(readResult, sheet, rowIndex, member)
                : new ReadCellResult(readResult.ColumnIndex, transformer.TransformStringValue(sheet, rowIndex, readResult), preserveFormatting);
        }

        return readResult;
    }

    /// <summary>
    /// Whether an invalid fallback uses the exception. Fixed value fallbacks ignore it, so it isn't created.
    /// </summary>
    internal static bool FallbackUsesException(IFallbackItem fallback)
    {
        return fallback.GetType() != typeof(FixedValueFallback);
    }

    /// <summary>
    /// Converts the result of <see cref="GetPropertyValue"/> to <typeparamref name="T"/>, throwing an
    /// <see cref="ExcelMappingException"/> rather than a <see cref="NullReferenceException"/> if the
    /// value is null and <typeparamref name="T"/> is a non-nullable value type.
    /// </summary>
    internal static T CastValue<T>(object? value, ExcelSheet sheet, int rowIndex, int columnIndex, MemberInfo? member)
    {
        // The error for this value has already been recorded (see ThrowFallback).
        if (ReferenceEquals(value, ThrowFallback.ErrorRecorded))
        {
            return default!;
        }

        if (value is null)
        {
            // default(T) is only non-null for non-nullable value types.
            if (default(T) is not null)
            {
                var memberDescription = member is null ? string.Empty : $" to member \"{member.Name}\"";
                throw new ExcelMappingException($"Cannot assign null{memberDescription} of type \"{typeof(T)}\". Consider adding a fallback or making the type nullable.", sheet, rowIndex, columnIndex)
                {
                    Member = member
                };
            }

            return default!;
        }

        return (T)value;
    }
}
