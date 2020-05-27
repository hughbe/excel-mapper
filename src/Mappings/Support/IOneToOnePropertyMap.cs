using System.Collections.Generic;

namespace ExcelMapper.Mappings.Support
{
    public interface IOneToOnePropertyMap
    {
        ISingleCellValueReader CellReader { get; set; }
        bool Optional { get; set; }
        IFallbackItem EmptyFallback { get; set; }
        IFallbackItem InvalidFallback { get; set; }

        IEnumerable<ICellValueTransformer> CellValueTransformers { get; }
        void AddCellValueTransformer(ICellValueTransformer transformer);

        IEnumerable<ICellValueMapper> CellValueMappers { get; }
        void AddCellValueMapper(ICellValueMapper item);
    }
}
