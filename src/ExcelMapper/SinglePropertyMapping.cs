using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Support;

namespace ExcelMapper
{
    public class SinglePropertyMapping : PropertyMapping, ISinglePropertyMapping
    {
        private List<IStringValueTransformer> _transformers = new List<IStringValueTransformer>();
        private List<ISinglePropertyMappingItem> _mappingItems = new List<ISinglePropertyMappingItem>();

        public Type Type { get; }

        public ISinglePropertyMapper Mapper { get; set; }

        public IEnumerable<IStringValueTransformer> StringValueTransformers => _transformers;
        public IEnumerable<ISinglePropertyMappingItem> MappingItems => _mappingItems;

        public void AddMappingItem(ISinglePropertyMappingItem item)
        {
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }

            _mappingItems.Add(item);
        }

        public void InsertMappingItem(int index, ISinglePropertyMappingItem item)
        {
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }

            _mappingItems.Insert(index, item);
        }

        public void AddStringValueTransformer(IStringValueTransformer transformer)
        {
            if (transformer == null)
            {
                throw new ArgumentNullException(nameof(transformer));
            }

            _transformers.Add(transformer);
        }

        public ISinglePropertyMappingItem EmptyFallback { get; set; }
        public ISinglePropertyMappingItem InvalidFallback { get; set; }

        public SinglePropertyMapping(MemberInfo member, Type type, EmptyValueStrategy emptyValueStrategy) : base(member)
        {
            Mapper = new ColumnPropertyMapper(member.Name);
            Type = type;
            AutoMapper.AutoMap(this, emptyValueStrategy);
        }

        public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            MapResult mapResult = Mapper.GetValue(sheet, rowIndex, reader);
            return GetPropertyValue(sheet, rowIndex, reader, mapResult);
        }

        internal object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MapResult mapResult)
        {
            if (mapResult.StringValue == null && EmptyFallback != null)
            {
                PropertyMappingResult fallbackResult = EmptyFallback.GetProperty(sheet, rowIndex, reader, mapResult);
                return fallbackResult.Value;
            }

            for (int i = 0; i < _transformers.Count; i++)
            {
                IStringValueTransformer transformer = _transformers[i];
                mapResult.StringValue = transformer.TransformStringValue(sheet, rowIndex, reader, mapResult);
            }

            var result = new PropertyMappingResult();
            for (int i = 0; i < _mappingItems.Count; i++)
            {
                ISinglePropertyMappingItem mappingItem = _mappingItems[i];

                PropertyMappingResult subResult = mappingItem.GetProperty(sheet, rowIndex, reader, mapResult);
                if (subResult.Type == PropertyMappingResultType.Success)
                {
                    return subResult.Value;
                }
                else if (subResult.Type != PropertyMappingResultType.Continue)
                {
                    result = subResult;
                }
            }

            if (result.Type != PropertyMappingResultType.Began && InvalidFallback != null)
            {
                PropertyMappingResult fallbackResult = InvalidFallback.GetProperty(sheet, rowIndex, reader, mapResult);
                return fallbackResult.Value;
            }

            return result.Value;
        }
    }
}
