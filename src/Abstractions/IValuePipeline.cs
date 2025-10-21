using System.Collections.Generic;

namespace ExcelMapper.Abstractions;

public interface IValuePipeline
{
    IList<ICellTransformer> Transformers { get; }
    IList<ICellMapper> Mappers { get; }
    IFallbackItem? EmptyFallback { get; set; }
    IFallbackItem? InvalidFallback { get; set; }
}
