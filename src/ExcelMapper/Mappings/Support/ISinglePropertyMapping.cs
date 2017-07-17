using System.Collections.Generic;

namespace ExcelMapper.Mappings.Support
{
    public interface ISinglePropertyMapping
    {
        ICellValueReader CellReader { get; set; }
        IFallbackItem EmptyFallback { get; set; }
        IFallbackItem InvalidFallback { get; set; }

        IEnumerable<ICellValueTransformer> CellValueTransformers { get; }
        void AddCellValueTransformer(ICellValueTransformer transformer);

        IEnumerable<ICellValueMapper> CellValueMappers { get; }
        void AddCellValueMapper(ICellValueMapper item);
    }
}
