using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper;

/// <summary>
/// Reads a single cell of an excel sheet and maps the value of the cell to the
/// type of the property or field.
/// </summary>
public class ValuePipeline<T> : ValuePipeline, IValuePipeline<T>
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
    internal static T Map(IValuePipeline pipeline, ExcelSheet sheet, int rowIndex, ReadCellResult readResult, bool preserveFormatting, MemberInfo? member)
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
    private sealed class Plan(int version, ICellConverter<T>[]? converters, ICellTransformer[] transformers, IFallbackItem? emptyFallback, IFallbackItem? invalidFallback)
    {
        public int Version { get; } = version;

        // The mappers as typed converters, or null if any of them isn't one (for example a custom
        // mapper), in which case values go through the pipeline as objects.
        public ICellConverter<T>[]? Converters { get; } = converters;

        public ICellTransformer[] Transformers { get; } = transformers;

        public IFallbackItem? EmptyFallback { get; } = emptyFallback;

        public IFallbackItem? InvalidFallback { get; } = invalidFallback;

        public bool InvalidFallbackUsesException { get; } = invalidFallback is not null && FallbackUsesException(invalidFallback);
    }

    private Plan? _plan;

    private Plan GetPlan()
    {
        var version = Version;
        var plan = _plan;
        if (plan is null || plan.Version != version)
        {
            plan = new Plan(version, CreateConverters((NonNullCollection<ICellMapper>)Mappers), [.. Transformers], EmptyFallback, InvalidFallback);
            _plan = plan;
        }

        return plan;
    }

    private static ICellConverter<T>[]? CreateConverters(NonNullCollection<ICellMapper> mappers)
    {
        var converters = new ICellConverter<T>[mappers.Count];
        for (int i = 0; i < converters.Length; i++)
        {
            if (mappers[i] is not ICellConverter<T> converter)
            {
                return null;
            }

            converters[i] = converter;
        }

        return converters;
    }

    /// <summary>
    /// Maps a cell through the typed converters, with the same rules as
    /// <see cref="ValuePipeline.GetPropertyValue"/>: transformers, the empty fallback, then the
    /// converters (a success stops, a success that can be replaced and an invalid value are kept
    /// until a later result replaces them), then the invalid fallback.
    /// </summary>
    private T Convert(Plan plan, ICellConverter<T>[] converters, ExcelSheet sheet, int rowIndex, ReadCellResult readResult, bool preserveFormatting, MemberInfo? member)
    {
        if (plan.Transformers.Length != 0)
        {
            readResult = ApplyTransformers(plan.Transformers, sheet, rowIndex, readResult, preserveFormatting);
        }

        if (plan.EmptyFallback is { } emptyFallback && readResult.IsEmpty())
        {
            return CastValue<T>(emptyFallback.PerformFallback(sheet, rowIndex, readResult, null, member), sheet, rowIndex, readResult.ColumnIndex, member)!;
        }

        var state = CellConversion.NotHandled;
        T result = default!;
        var invalidIndex = -1;
        for (int i = 0; i < converters.Length; i++)
        {
            switch (converters[i].Convert(readResult, out var value))
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
                    break;
            }
        }

        if (state == CellConversion.SuccessUnlessReplaced)
        {
            return result;
        }

        if (plan.InvalidFallback is { } invalidFallback)
        {
            // Create the exception with the mapper's own Map, only if the fallback uses it.
            var exception = state == CellConversion.Invalid && plan.InvalidFallbackUsesException
                ? converters[invalidIndex].Map(readResult).Exception
                : null;
            return CastValue<T>(invalidFallback.PerformFallback(sheet, rowIndex, readResult, exception, member), sheet, rowIndex, readResult.ColumnIndex, member)!;
        }

        return CastValue<T>(null, sheet, rowIndex, readResult.ColumnIndex, member)!;
    }
}
