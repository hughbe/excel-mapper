using System.Collections.Generic;

namespace ExcelMapper.Mappings.Support
{
    public interface ISinglePropertyMapping
    {
        ISingleValueReader Reader { get; set; }
        IFallbackItem EmptyFallback { get; set; }
        IFallbackItem InvalidFallback { get; set; }

        IEnumerable<IStringValueTransformer> StringValueTransformers { get; }
        void AddStringValueTransformer(IStringValueTransformer transformer);

        IEnumerable<ISinglePropertyMappingItem> MappingItems { get; }
        void AddMappingItem(ISinglePropertyMappingItem item);
    }
}
