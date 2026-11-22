using System.Reflection;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Transformers;

namespace ExcelMapper;

/// <summary>
/// How the cells of a member of type <typeparamref name="T"/> are converted: see
/// <see cref="ValuePipeline"/>.
/// </summary>
/// <typeparam name="T">The type of the member, or of the elements or values of a collection.</typeparam>
public abstract class ValuePipeline<T> : ValuePipeline
{
    private protected ValuePipeline()
    {
    }

    /// <summary>
    /// Gets the mappers, which are tried in order until one converts the value.
    /// </summary>
    internal IList<ICellMapper<T>> Mappers { get; } = new NonNullCollection<ICellMapper<T>>();

    /// <summary>
    /// Gets or sets the value of empty cells. Null means empty cells are converted by the mappers.
    /// </summary>
    internal ICellFallback<T>? EmptyFallback
    {
        get => _emptyFallback;
        set
        {
            _emptyFallback = value;
            _fallbackVersion++;
        }
    }

    private ICellFallback<T>? _emptyFallback;

    /// <summary>
    /// Gets or sets the value of cells that no mapper converts. Null means the default value.
    /// </summary>
    internal ICellFallback<T>? InvalidFallback
    {
        get => _invalidFallback;
        set
        {
            _invalidFallback = value;
            _fallbackVersion++;
        }
    }

    private ICellFallback<T>? _invalidFallback;

    private int _fallbackVersion;

    internal override IEnumerable<object> MapperObjects => Mappers;

    internal override int MapperCount => Mappers.Count;

    internal override void AddMapper(object mapper) => Mappers.Add((ICellMapper<T>)mapper);

    internal override void SetFixedFallback(object? value, bool empty, bool invalid)
    {
        var fallback = new FixedValueFallback<T>(value);
        if (empty)
        {
            EmptyFallback = fallback;
        }

        if (invalid)
        {
            InvalidFallback = fallback;
        }
    }

    internal override void SetThrowingFallback(bool empty, bool invalid)
    {
        if (empty)
        {
            EmptyFallback = ThrowFallback<T>.Instance;
        }

        if (invalid)
        {
            InvalidFallback = ThrowFallback<T>.Instance;
        }
    }

    private int Version => TransformersVersion + ((NonNullCollection<ICellMapper<T>>)Mappers).Version + _fallbackVersion;

    /// <summary>
    /// Maps a cell to a <typeparamref name="T"/>, without boxing: the transformers, then the empty
    /// fallback for an empty cell, then the mappers in order (a success stops; a success that can
    /// be replaced and an invalid value are kept until a later result replaces them), then the
    /// invalid fallback.
    /// </summary>
    internal T Map(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, bool preserveFormatting, MemberInfo? member)
    {
        var plan = GetPlan();

        // Cells are parsed and formatted with the importer's culture, if set, rather than the
        // thread's: mappers use it when they have no format provider of their own.
        var culture = sheet.Culture;
        readResult.Culture = culture;

        // Use indexed loops: enumerating through an interface allocates for every cell.
        var transformers = plan.Transformers;
        for (int i = 0; i < transformers.Length; i++)
        {
            readResult = transformers[i].Transform(readResult, sheet, rowIndex, member, preserveFormatting);
            readResult.Culture = culture;
        }

        if (plan.EmptyFallback is { } emptyFallback && readResult.IsEmpty())
        {
            return emptyFallback.Fallback(new CellValue(readResult), new CellContext(sheet, rowIndex, readResult.ColumnIndex, member), null);
        }

        var state = CellConversion.NotHandled;
        T result = default!;
        var invalidIndex = -1;
        object? invalidError = null;
        var mappers = plan.Mappers;
        var cellConverters = plan.CellConverters;
        for (int i = 0; i < mappers.Length; i++)
        {
            CellConversion conversion;
            T value;
            if (mappers[i] is { } mapper)
            {
                conversion = mapper.Convert(readResult, out value);
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
                    if (mappers[i] is not null)
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
            // Create the exception only if the fallback uses it: from the mapper, or from the error
            // an ICellConverter<T> returned.
            var exception = state == CellConversion.Invalid && plan.InvalidFallbackUsesException
                ? mappers[invalidIndex] is { } invalidMapper
                    ? invalidMapper.CreateException(readResult)
                    : CellConversionResult<T>.CreateException(invalidError)
                : null;
            return invalidFallback.Fallback(new CellValue(readResult), new CellContext(sheet, rowIndex, readResult.ColumnIndex, member), exception);
        }

        return CastValue<T>(null, sheet, rowIndex, readResult.ColumnIndex, member)!;
    }

    // An ICellConverter<T> gets the cell's context, and its error is kept from the conversion: it
    // can't be asked for an exception afterwards without converting the value again.
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

    // A snapshot of the pipeline, made when it's first used and again after it changes, so that
    // mapping a cell reads arrays rather than lists through interfaces.
    private sealed class Plan(int version, ICellMapper<T>?[] mappers, ICellConverter<T>?[]? cellConverters, IPipelineTransformer[] transformers, ICellFallback<T>? emptyFallback, ICellFallback<T>? invalidFallback)
    {
        public int Version { get; } = version;

        // The mappers. The mapper for an ICellConverter<T> is null here, and its converter is in
        // CellConverters.
        public ICellMapper<T>?[] Mappers { get; } = mappers;

        // The ICellConverter<T> of each mapper that has one (see CellConverterMapper<T>), or null
        // if no mapper has one.
        public ICellConverter<T>?[]? CellConverters { get; } = cellConverters;

        public IPipelineTransformer[] Transformers { get; } = transformers;

        public ICellFallback<T>? EmptyFallback { get; } = emptyFallback;

        public ICellFallback<T>? InvalidFallback { get; } = invalidFallback;

        // Fixed values don't use the exception, so it isn't created for them.
        public bool InvalidFallbackUsesException { get; } = invalidFallback is not null and not FixedValueFallback<T>;
    }

    private Plan? _plan;

    private Plan GetPlan()
    {
        var version = Version;
        var plan = _plan;
        if (plan is null || plan.Version != version)
        {
            var mappers = new ICellMapper<T>?[Mappers.Count];
            ICellConverter<T>?[]? cellConverters = null;
            for (int i = 0; i < mappers.Length; i++)
            {
                if (Mappers[i] is CellConverterMapper<T> cellConverterMapper)
                {
                    cellConverters ??= new ICellConverter<T>?[mappers.Length];
                    cellConverters[i] = cellConverterMapper.Converter;
                }
                else
                {
                    mappers[i] = Mappers[i];
                }
            }

            plan = new Plan(version, mappers, cellConverters, [.. Transformers], EmptyFallback, InvalidFallback);
            _plan = plan;
        }

        return plan;
    }
}
