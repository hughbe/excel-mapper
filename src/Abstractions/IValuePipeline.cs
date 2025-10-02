using System.Collections.Generic;

namespace ExcelMapper.Abstractions;

public interface IValuePipeline
{
    IReadOnlyList<ICellTransformer> CellValueTransformers { get; }
    IReadOnlyList<ICellMapper> CellValueMappers { get; }
    void AddCellValueMapper(ICellMapper mapper);
    void RemoveCellValueMapper(int index);
    void AddCellValueTransformer(ICellTransformer transformer);
    IFallbackItem? EmptyFallback { get; set; }
    IFallbackItem? InvalidFallback { get; set; }
}
