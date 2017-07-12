using System;
using System.Globalization;

namespace ExcelMapper.Mappings.Mappers
{
    public class DateTimeMapper : IStringValueMapper
    {
        private string[] _formats = new string[] { "G" };

        /// <summary>
        /// Defaults to "G" - the default Excel format.
        /// </summary>
        public string[] Formats
        {
            get => _formats;
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }

                if (value.Length == 0)
                {
                    throw new ArgumentException("Formats cannot be empty.", nameof(value));
                }

                _formats = value;
            }
        }

        public IFormatProvider Provider { get; set; }

        public DateTimeStyles Style { get; set; }

        public PropertyMappingResultType GetProperty(ReadResult readResult, ref object value)
        {
            if (!DateTime.TryParseExact(readResult.StringValue, Formats, Provider, Style, out DateTime result))
            {
                return PropertyMappingResultType.Invalid;
            }

            value = result;
            return PropertyMappingResultType.Success;
        }
    }
}
