using System.Reflection;
using ExcelMapper.Transformers;

namespace ExcelMapper;

/// <summary>
/// How the cells of a member are converted to its type: the transformers, converters and
/// fallbacks set with the value methods of <see cref="ValuePipeline{T, TSelf}"/> and
/// <see cref="ValuePipelineExtensions"/>. A member's map (<see cref="ValueMap{T}"/>) has one, and
/// so do the elements or values of collections and dictionaries (<see cref="ElementMap{T}"/>).
/// </summary>
internal abstract class CellPipeline
{
    // Pipelines are typed: see CellPipeline<T>.
    private protected CellPipeline()
    {
    }

    /// <summary>
    /// Gets the transformers, which change each cell, in order, before it's converted.
    /// </summary>
    internal NonNullCollection<ICellTransformer> Transformers { get; } = [];

    /// <summary>
    /// Gets whether the pipeline has been used to read rows, so it can't be changed.
    /// </summary>
    internal bool IsFrozen { get; private set; }

    /// <summary>
    /// Throws if the pipeline has been used to read rows.
    /// </summary>
    internal void ThrowIfFrozen() => ThrowHelpers.ThrowIfFrozen(IsFrozen);

    /// <summary>
    /// Stops the pipeline changing, the first time rows are read with it (see <see cref="IMap.Freeze"/>).
    /// </summary>
    internal virtual void Freeze()
    {
        IsFrozen = true;
        Transformers.Freeze();
        Options.Freeze();
    }

    /// <summary>
    /// Gets how the built-in mappers parse text: formats, format provider, number style and so on.
    /// </summary>
    internal ConversionOptions Options { get; } = new();

    /// <summary>
    /// Gets whether the pipeline has any mappers.
    /// </summary>
    internal abstract bool HasMappers { get; }

    /// <summary>
    /// Gives empty cells, invalid cells or both a fixed value, for the untyped fluent methods
    /// (<c>WithEmptyFallback(object)</c> and so on).
    /// </summary>
    internal abstract void SetFixedFallback(object? value, bool empty, bool invalid);

    /// <summary>
    /// Makes empty cells, invalid cells or both errors.
    /// </summary>
    internal abstract void SetThrowingFallback(bool empty, bool invalid);

    /// <summary>
    /// Converts a fallback's value to <typeparamref name="T"/>, throwing an
    /// <see cref="ExcelMappingException"/> rather than a <see cref="NullReferenceException"/> if the
    /// value is null and <typeparamref name="T"/> is a non-nullable value type.
    /// </summary>
    internal static T CastValue<T>(object? value, SheetView sheet, int rowIndex, int columnIndex, MemberInfo? member)
    {
        if (value is null)
        {
            // default(T) is only non-null for non-nullable value types.
            if (default(T) is not null)
            {
                var memberDescription = member is null ? string.Empty : $" to member \"{member.Name}\"";
                throw new ExcelMappingException($"Cannot assign null{memberDescription} of type {TypeNames.Get(typeof(T))}. Consider adding a fallback or making the type nullable.", sheet, rowIndex, columnIndex)
                {
                    Member = member
                };
            }

            return default!;
        }

        return (T)value;
    }
}
