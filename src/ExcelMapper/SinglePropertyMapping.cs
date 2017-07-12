using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Readers;
using ExcelMapper.Mappings.Support;

namespace ExcelMapper
{
    public class SinglePropertyMapping : PropertyMapping, ISinglePropertyMapping
    {
        private List<IStringValueTransformer> _transformers = new List<IStringValueTransformer>();
        private List<ISinglePropertyMappingItem> _mappingItems = new List<ISinglePropertyMappingItem>();

        public Type Type { get; }

        public ISingleValueReader Reader { get; set; }

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

        public void RemoveMappingItem(int index) => _mappingItems.RemoveAt(index);

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
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            if (!Enum.IsDefined(typeof(EmptyValueStrategy), emptyValueStrategy))
            {
                throw new ArgumentException($"Invalid EmptyValueStategy \"{emptyValueStrategy}\".", nameof(emptyValueStrategy));
            }

            Reader = new ColumnNameReader(member.Name);
            Type = type;
            AutoMapper.AutoMap(this, emptyValueStrategy);
        }

        public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            ReadResult mapResult = Reader.GetValue(sheet, rowIndex, reader);
            return GetPropertyValue(sheet, rowIndex, reader, mapResult);
        }

        internal object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, ReadResult mapResult)
        {
            for (int i = 0; i < _transformers.Count; i++)
            {
                IStringValueTransformer transformer = _transformers[i];
                mapResult = new ReadResult(mapResult.ColumnIndex, transformer.TransformStringValue(sheet, rowIndex, reader, mapResult));
            }

            if (mapResult.StringValue == null && EmptyFallback != null)
            {
                PropertyMappingResult fallbackResult = EmptyFallback.GetProperty(sheet, rowIndex, reader, mapResult);
                return fallbackResult.Value;
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
