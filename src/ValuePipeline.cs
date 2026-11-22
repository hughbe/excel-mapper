using System.Reflection;
using ExcelMapper.Transformers;

namespace ExcelMapper;

/// <summary>
/// How the cells of a member are converted to its type: the transformers, converters and
/// fallbacks set with the value methods in <see cref="ValuePipelineExtensions"/>. A member's map
/// (<see cref="OneToOneMap{T}"/>) is a pipeline, and collections and dictionaries have one for
/// their elements or values.
/// </summary>
public abstract class ValuePipeline
{
    // Pipelines are typed: see ValuePipeline<T>.
    private protected ValuePipeline()
    {
    }

    /// <summary>
    /// Gets the transformers, which change each cell, in order, before it's converted.
    /// </summary>
    internal IList<IPipelineTransformer> Transformers { get; } = new NonNullCollection<IPipelineTransformer>();

    private protected int TransformersVersion => ((NonNullCollection<IPipelineTransformer>)Transformers).Version;

    /// <summary>
    /// Gets the mappers, for code that configures them without knowing the pipeline's type (such
    /// as <see cref="ExcelFormatsAttribute"/>).
    /// </summary>
    internal abstract IEnumerable<object> MapperObjects { get; }

    /// <summary>
    /// Gets the number of mappers.
    /// </summary>
    internal abstract int MapperCount { get; }

    /// <summary>
    /// Adds a mapper, which must be one for the pipeline's type.
    /// </summary>
    internal abstract void AddMapper(object mapper);

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
    internal static T CastValue<T>(object? value, ExcelSheet sheet, int rowIndex, int columnIndex, MemberInfo? member)
    {
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
