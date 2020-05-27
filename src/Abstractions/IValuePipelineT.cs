using System.Collections.Generic;

namespace ExcelMapper.Abstractions
{
    public interface IValuePipeline
    {
        IEnumerable<ICellValueTransformer> CellValueTransformers { get; }
        IEnumerable<ICellValueMapper> CellValueMappers { get; }
        void AddCellValueMapper(ICellValueMapper mapper);
        void RemoveCellValueMapper(int index);
        void AddCellValueTransformer(ICellValueTransformer transformer);
        IFallbackItem EmptyFallback { get; set; }
        IFallbackItem InvalidFallback { get; set; }
    }
}
