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
        private List<IStringValueMapper> _mappingItems = new List<IStringValueMapper>();

        public Type Type { get; }

        public ISingleValueReader Reader { get; set; }

        public IEnumerable<IStringValueTransformer> StringValueTransformers => _transformers;
        public IEnumerable<IStringValueMapper> MappingItems => _mappingItems;

        public void AddMappingItem(IStringValueMapper item)
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

        public IFallbackItem EmptyFallback { get; set; }
        public IFallbackItem InvalidFallback { get; set; }

        public SinglePropertyMapping(MemberInfo member, Type type) : base(member)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            Reader = new ColumnNameReader(member.Name);
            Type = type;
        }

        public override object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            ReadResult readResult = Reader.GetValue(sheet, rowIndex, reader);
            return GetPropertyValue(sheet, rowIndex, reader, readResult);
        }

        internal object GetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, ReadResult readResult)
        {
            for (int i = 0; i < _transformers.Count; i++)
            {
                IStringValueTransformer transformer = _transformers[i];
                readResult = new ReadResult(readResult.ColumnIndex, transformer.TransformStringValue(sheet, rowIndex, readResult));
            }

            if (readResult.StringValue == null && EmptyFallback != null)
            {
                return EmptyFallback.PerformFallback(sheet, rowIndex, readResult);
            }

            PropertyMappingResultType resultType = PropertyMappingResultType.Success;
            object value = null;

            for (int i = 0; i < _mappingItems.Count; i++)
            {
                IStringValueMapper mappingItem = _mappingItems[i];

                PropertyMappingResultType newResultType  = mappingItem.GetProperty(readResult, ref value);
                if (newResultType == PropertyMappingResultType.Success)
                {
                    return value;
                }
                else if (newResultType != PropertyMappingResultType.Continue)
                {
                    resultType = newResultType;
                }
            }

            if (resultType != PropertyMappingResultType.Success && resultType != PropertyMappingResultType.SuccessIfNoOtherSuccess && InvalidFallback != null)
            {
                return InvalidFallback.PerformFallback(sheet, rowIndex, readResult);
            }

            return value;
        }
    }
}
