using System.Reflection;
using ExcelMapper.Fallbacks;

namespace ExcelMapper;

/// <summary>
/// How the cells of a member of type <typeparamref name="T"/> are converted: see
/// <see cref="CellPipeline"/>.
/// </summary>
/// <typeparam name="T">The type of the member, or of the elements or values of a collection.</typeparam>
internal sealed class CellPipeline<T> : CellPipeline
{
    internal CellPipeline()
    {
    }

    /// <summary>
    /// Gets the mappers, which are tried in order until one converts the value.
    /// </summary>
    internal NonNullCollection<ICellMapper<T>> Mappers { get; } = [];

    /// <summary>
    /// Gets or sets the value of empty cells. Null means empty cells are converted by the mappers.
    /// </summary>
    internal ICellFallback<T>? EmptyFallback
    {
        get => _emptyFallback;
        set
        {
            ThrowIfFrozen();
            _emptyFallback = value;
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
            ThrowIfFrozen();
            _invalidFallback = value;
        }
    }

    private ICellFallback<T>? _invalidFallback;

    internal override bool HasMappers => Mappers.Count > 0;

    internal override void Freeze()
    {
        if (IsFrozen)
        {
            return;
        }

        base.Freeze();
        Mappers.Freeze();
        _plan = CreatePlan();
    }

    /// <summary>
    /// Gives empty cells, invalid cells or both a fixed value (<c>WithEmptyFallback(value)</c> and so on).
    /// </summary>
    internal void SetFixedFallback(T value, bool empty, bool invalid)
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

    /// <summary>
    /// Maps a cell to a <typeparamref name="T"/>, without boxing: the transformers, then the empty
    /// fallback for an empty cell, then the mappers in order (a success stops; a success that can
    /// be replaced and an invalid value are kept until a later result replaces them), then the
    /// invalid fallback.
    /// </summary>
    internal T Map(SheetView sheet, int rowIndex, ReadCellResult readResult, MemberInfo? member, ErrorTarget? errors)
    {
        var plan = GetPlan();

        // Cells are parsed and formatted with the importer's culture, if set, rather than the
        // thread's: mappers use it when they have no format provider of their own.
        var culture = sheet.Culture;
        readResult.Culture = culture;
        var context = new CellContext(sheet, rowIndex, readResult.ColumnIndex, member, Options);

        // Use indexed loops: enumerating through an interface allocates for every cell.
        var transformers = plan.Transformers;
        for (int i = 0; i < transformers.Length; i++)
        {
            readResult = transformers[i].Transform(new CellValue(readResult), context).ToReadCellResult(readResult.ColumnIndex);
            readResult.Culture = culture;
        }

        if (plan.EmptyFallback is { } emptyFallback && readResult.IsEmpty())
        {
            // A collected error is reported without creating its exception (see ErrorTarget).
            if (errors is not null && emptyFallback is ThrowFallback<T>)
            {
                errors.AddCell(ExcelMappingErrorKind.EmptyValue, readResult, in context, typeof(T), null, null, ThrowFallback<T>.EmptyHint);
                return default!;
            }

            return emptyFallback.Fallback(new CellValue(readResult), context, null);
        }

        var state = CellConversion.NotHandled;
        T result = default!;
        var invalidIndex = -1;
        object? invalidError = null;
        var mappers = plan.Mappers;
        for (int i = 0; i < mappers.Length; i++)
        {
            object? error = null;
            switch (mappers[i].Convert(readResult, in context, out var value, ref error))
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
                    invalidError = error;
                    break;
            }
        }

        if (state == CellConversion.SuccessUnlessReplaced)
        {
            return result;
        }

        if (plan.InvalidFallback is { } invalidFallback)
        {
            // A collected error is reported without creating its exception: the mapper that
            // found the cell invalid creates it if the target asks (see ErrorTarget).
            if (errors is not null && invalidFallback is ThrowFallback<T>)
            {
                var mapper = state == CellConversion.Invalid ? mappers[invalidIndex] : null;
                errors.AddCell(ExcelMappingErrorKind.InvalidValue, readResult, in context, typeof(T), mapper, invalidError, ThrowFallback<T>.EmptyHint);
                return default!;
            }

            // Create the exception only if the fallback uses it.
            var exception = state == CellConversion.Invalid && plan.InvalidFallbackUsesException
                ? mappers[invalidIndex].CreateException(readResult, in context, invalidError)
                : null;
            return invalidFallback.Fallback(new CellValue(readResult), context, exception);
        }

        return CastValue<T>(null, sheet, rowIndex, readResult.ColumnIndex, member)!;
    }

    // A snapshot of the pipeline, made when it's frozen, so that mapping a cell reads arrays
    // rather than lists through interfaces.
    private sealed class Plan(ICellMapper<T>[] mappers, ICellTransformer[] transformers, ICellFallback<T>? emptyFallback, ICellFallback<T>? invalidFallback)
    {
        public ICellMapper<T>[] Mappers { get; } = mappers;

        public ICellTransformer[] Transformers { get; } = transformers;

        public ICellFallback<T>? EmptyFallback { get; } = emptyFallback;

        public ICellFallback<T>? InvalidFallback { get; } = invalidFallback;

        // Fixed values don't use the exception, so it isn't created for them.
        public bool InvalidFallbackUsesException { get; } = invalidFallback is not null and not FixedValueFallback<T>;
    }

    private Plan? _plan;

    // Maps are frozen before rows are read, so the plan is made then. A pipeline is also frozen
    // the first time it maps a cell, in case it's used on its own. (Another thread may be freezing
    // it at the same time, so the plan may not be set yet: plans of a frozen pipeline are equal.)
    private Plan GetPlan() => _plan ?? FreezeAndGetPlan();

    private Plan FreezeAndGetPlan()
    {
        Freeze();
        return _plan ?? CreatePlan();
    }

    private Plan CreatePlan() => new([.. Mappers], [.. Transformers], EmptyFallback, InvalidFallback);
}
