using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;

namespace ExcelMapper;

/// <summary>
/// Reads a single cell of an excel sheet and maps the value of the cell to the
/// type of the property or field.
/// </summary>
public class ValuePipeline<T> : ValuePipeline
{
    /// <summary>
    /// Maps a cell to a <typeparamref name="T"/>: through the typed converters if every mapper is
    /// one, without boxing, and otherwise through the pipeline as objects.
    /// </summary>
    internal T Map(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, bool preserveFormatting, MemberInfo? member)
    {
        var plan = GetPlan();
        if (plan.Converters is { } converters)
        {
            return Convert(plan, converters, sheet, rowIndex, readResult, preserveFormatting, member);
        }

        var result = GetPropertyValue(this, sheet, rowIndex, readResult, preserveFormatting, member);
        return CastValue<T>(result, sheet, rowIndex, readResult.ColumnIndex, member)!;
    }

    /// <summary>
    /// Maps a cell to a <typeparamref name="T"/> with any pipeline: <see cref="Map"/> for a
    /// <see cref="ValuePipeline{T}"/>, and otherwise through the pipeline as objects.
    /// </summary>
    internal static T Map(ValuePipeline pipeline, ExcelSheet sheet, int rowIndex, ReadCellResult readResult, bool preserveFormatting, MemberInfo? member)
    {
        if (pipeline is ValuePipeline<T> typedPipeline)
        {
            return typedPipeline.Map(sheet, rowIndex, readResult, preserveFormatting, member);
        }

        var result = GetPropertyValue(pipeline, sheet, rowIndex, readResult, preserveFormatting, member);
        return CastValue<T>(result, sheet, rowIndex, readResult.ColumnIndex, member)!;
    }

    /// <summary>
    /// The pipeline as arrays and fields, so that mapping a cell doesn't go through the pipeline's
    /// interfaces. Rebuilt when the pipeline changes.
    /// </summary>
    private sealed class Plan(int version, ITypedCellMapper<T>?[]? converters, ICellConverter<T>?[]? cellConverters, Abstractions.ICellTransformer[] transformers, IFallbackItem? emptyFallback, IFallbackItem? invalidFallback)
    {
        public int Version { get; } = version;

        // The mappers as typed converters, or null if any of them isn't one (for example a custom
        // mapper), in which case values go through the pipeline as objects. A mapper for an
        // ICellConverter<T> is null here, and its converter is in CellConverters.
        public ITypedCellMapper<T>?[]? Converters { get; } = converters;

        // The ICellConverter<T> of each mapper that has one (see CellConverterMapper<T>), or null
        // if no mapper has one.
        public ICellConverter<T>?[]? CellConverters { get; } = cellConverters;

        public Abstractions.ICellTransformer[] Transformers { get; } = transformers;

        public IFallbackItem? EmptyFallback { get; } = emptyFallback;

        public IFallbackItem? InvalidFallback { get; } = invalidFallback;

        public bool InvalidFallbackUsesException { get; } = invalidFallback is not null && FallbackUsesException(invalidFallback);

        // The ICellFallback<T> of each fallback that has one, called without boxing.
        public ICellFallback<T>? TypedEmptyFallback { get; } = (emptyFallback as CellFallbackItem<T>)?.Fallback;

        public ICellFallback<T>? TypedInvalidFallback { get; } = (invalidFallback as CellFallbackItem<T>)?.Fallback;
    }

    private Plan? _plan;

    private Plan GetPlan()
    {
        var version = Version;
        var plan = _plan;
        if (plan is null || plan.Version != version)
        {
            var converters = CreateConverters((NonNullCollection<ICellMapper>)Mappers, out var cellConverters);
            plan = new Plan(version, converters, cellConverters, [.. Transformers], EmptyFallback, InvalidFallback);
            _plan = plan;
        }

        return plan;
    }

    private static ITypedCellMapper<T>?[]? CreateConverters(NonNullCollection<ICellMapper> mappers, out ICellConverter<T>?[]? cellConverters)
    {
        cellConverters = null;
        var converters = new ITypedCellMapper<T>?[mappers.Count];
        for (int i = 0; i < converters.Length; i++)
        {
            if (mappers[i] is CellConverterMapper<T> cellConverterMapper)
            {
                cellConverters ??= new ICellConverter<T>?[converters.Length];
                cellConverters[i] = cellConverterMapper.Converter;
            }
            else if (mappers[i] is ITypedCellMapper<T> converter)
            {
                converters[i] = converter;
            }
            else
            {
                cellConverters = null;
                return null;
            }
        }

        return converters;
    }

    /// <summary>
    /// Maps a cell through the typed converters, with the same rules as
    /// <see cref="ValuePipeline.GetPropertyValue"/>: transformers, the empty fallback, then the
    /// converters (a success stops, a success that can be replaced and an invalid value are kept
    /// until a later result replaces them), then the invalid fallback.
    /// </summary>
    private T Convert(Plan plan, ITypedCellMapper<T>?[] converters, ExcelSheet sheet, int rowIndex, ReadCellResult readResult, bool preserveFormatting, MemberInfo? member)
    {
        if (plan.Transformers.Length != 0)
        {
            readResult = ApplyTransformers(plan.Transformers, sheet, rowIndex, readResult, preserveFormatting, member);
        }

        if (plan.EmptyFallback is { } emptyFallback && readResult.IsEmpty())
        {
            if (plan.TypedEmptyFallback is { } typedEmptyFallback)
            {
                return typedEmptyFallback.Fallback(new CellValue(readResult), new CellContext(sheet, rowIndex, readResult.ColumnIndex, member), null);
            }

            return CastValue<T>(emptyFallback.PerformFallback(sheet, rowIndex, readResult, null, member), sheet, rowIndex, readResult.ColumnIndex, member)!;
        }

        var state = CellConversion.NotHandled;
        T result = default!;
        var invalidIndex = -1;
        object? invalidError = null;
        var cellConverters = plan.CellConverters;
        for (int i = 0; i < converters.Length; i++)
        {
            CellConversion conversion;
            T value;
            if (converters[i] is { } converter)
            {
                conversion = converter.Convert(readResult, out value);
            }
            else
            {
                conversion = ConvertWithCellConverter(cellConverters![i]!, sheet, rowIndex, readResult, member, out value, ref invalidError);
            }

            switch (conversion)
            {
                case CellConversion.Success:
                    return value;
                case CellConversion.SuccessUnlessReplaced:
                    state = CellConversion.SuccessUnlessReplaced;
                    result = value;
                    break;
                case CellConversion.Invalid:
                    state = CellConversion.Invalid;
                    invalidIndex = i;
                    if (converters[i] is not null)
                    {
                        invalidError = null;
                    }

                    break;
            }
        }

        if (state == CellConversion.SuccessUnlessReplaced)
        {
            return result;
        }

        if (plan.InvalidFallback is { } invalidFallback)
        {
            // Create the exception only if the fallback uses it: with the built-in mapper's own Map, or
            // from the error an ICellConverter<T> returned.
            var exception = state == CellConversion.Invalid && plan.InvalidFallbackUsesException
                ? converters[invalidIndex] is { } invalidConverter
                    ? invalidConverter.Map(readResult).Exception
                    : CellConversionResult<T>.CreateException(invalidError)
                : null;
            if (plan.TypedInvalidFallback is { } typedInvalidFallback)
            {
                return typedInvalidFallback.Fallback(new CellValue(readResult), new CellContext(sheet, rowIndex, readResult.ColumnIndex, member), exception);
            }

            return CastValue<T>(invalidFallback.PerformFallback(sheet, rowIndex, readResult, exception, member), sheet, rowIndex, readResult.ColumnIndex, member)!;
        }

        return CastValue<T>(null, sheet, rowIndex, readResult.ColumnIndex, member)!;
    }

    private static CellConversion ConvertWithCellConverter(ICellConverter<T> converter, ExcelSheet sheet, int rowIndex, ReadCellResult readResult, MemberInfo? member, out T value, ref object? invalidError)
    {
        var result = converter.Convert(new CellValue(readResult), new CellContext(sheet, rowIndex, readResult.ColumnIndex, member));
        value = result.Value;
        switch (result.Status)
        {
            case CellConversionStatus.Success:
                return CellConversion.Success;
            case CellConversionStatus.Invalid:
                invalidError = result.Error;
                return CellConversion.Invalid;
            default:
                return CellConversion.NotHandled;
        }
    }
}
