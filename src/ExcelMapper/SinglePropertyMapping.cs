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
            int columnIndex = Mapper.GetColumnIndex(sheet, rowIndex, reader);
            return GetPropertyValue(sheet, rowIndex, reader, columnIndex);
        }

        internal object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex)
        {
            string stringValue = reader.GetString(columnIndex);
            if (stringValue == null && EmptyFallback != null)
            {
                PropertyMappingResult fallbackResult = EmptyFallback.GetProperty(sheet, rowIndex, reader, columnIndex, stringValue);
                return fallbackResult.Value;
            }

            for (int i = 0; i < _transformers.Count; i++)
            {
                IStringValueTransformer transformer = _transformers[i];
                stringValue = transformer.TransformStringValue(sheet, rowIndex, reader, columnIndex, stringValue);
            }

            var result = new PropertyMappingResult();
            for (int i = 0; i < _mappingItems.Count; i++)
            {
                ISinglePropertyMappingItem mappingItem = _mappingItems[i];

                result = mappingItem.GetProperty(sheet, rowIndex, reader, columnIndex, stringValue);
                if (result.Type == PropertyMappingResultType.Invalid && InvalidFallback != null)
                {
                    PropertyMappingResult fallbackResult = InvalidFallback.GetProperty(sheet, rowIndex, reader, columnIndex, stringValue);
                    return fallbackResult.Value;
                }
            }

            return result.Value;
        }
    }
}
