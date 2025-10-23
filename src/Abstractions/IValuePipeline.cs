namespace ExcelMapper.Abstractions;

/// <summary>
/// Pipeline for processing cell values through transformers and mappers.
/// </summary>
public interface IValuePipeline
{
    /// <summary>
    /// The list of transformers in the pipeline.
    /// </summary>
    IList<ICellTransformer> Transformers { get; }

    /// <summary>
    /// The list of mappers in the pipeline.
    /// </summary>
    IList<ICellMapper> Mappers { get; }

    /// <summary>
    /// The fallback item for empty values.
    /// </summary>
    IFallbackItem? EmptyFallback { get; set; }

    /// <summary>
    /// The fallback item for invalid values.
    /// </summary>
    IFallbackItem? InvalidFallback { get; set; }
}
