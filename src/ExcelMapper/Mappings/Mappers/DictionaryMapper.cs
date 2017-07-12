using System;
using System.Collections.Generic;

namespace ExcelMapper.Mappings.Mappers
{
    public class DictionaryMapper<T> : IStringValueMapper
    {
        public IReadOnlyDictionary<string, T> MappingDictionary { get; }

        public DictionaryMapper(IDictionary<string, T> mappingDictionary, IEqualityComparer<string> comparer)
        {
            if (mappingDictionary == null)
            {
                throw new ArgumentNullException(nameof(mappingDictionary));
            }

            MappingDictionary = new Dictionary<string, T>(mappingDictionary, comparer);
        }

        public PropertyMappingResultType GetProperty(ReadResult readResult, ref object value)
        {
            // If we didn't find anything, keep going. This is not necessarily a fatal error.
            if (!MappingDictionary.TryGetValue(readResult.StringValue, out T result))
            {
                return PropertyMappingResultType.Continue;
            }

            value = result;
            return PropertyMappingResultType.Success;
        }
    }
}
