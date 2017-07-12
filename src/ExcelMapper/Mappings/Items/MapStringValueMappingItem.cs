using System;
using System.Collections.Generic;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    public class MapStringValueMappingItem<T> : ISinglePropertyMappingItem
    {
        public IReadOnlyDictionary<string, T> MappingDictionary { get; }

        public MapStringValueMappingItem(IDictionary<string, T> mappingDictionary, IEqualityComparer<string> comparer)
        {
            if (mappingDictionary == null)
            {
                throw new ArgumentNullException(nameof(mappingDictionary));
            }

            MappingDictionary = new Dictionary<string, T>(mappingDictionary, comparer);
        }

        public PropertyMappingResult GetProperty(ReadResult mapResult)
        {
            if (!MappingDictionary.TryGetValue(mapResult.StringValue, out T result))
            {
                return PropertyMappingResult.Continue();
            }

            return PropertyMappingResult.Success(result);
        }
    }
}
