using System.Collections.Generic;

namespace ExcelMapper.Mappings.Support
{
    public interface ISinglePropertyMapping
    {
        ISinglePropertyMapper Mapper { get; set; }
        ISinglePropertyMappingItem EmptyFallback { get; set; }
        ISinglePropertyMappingItem InvalidFallback { get; set; }

        IEnumerable<IStringValueTransformer> StringValueTransformers { get; }
        void AddStringValueTransformer(IStringValueTransformer transformer);

        IEnumerable<ISinglePropertyMappingItem> MappingItems { get; }
        void AddMappingItem(ISinglePropertyMappingItem mappingItem);
    }
}
