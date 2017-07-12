using System;
using System.Globalization;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    public class ParseAsDateTimeMappingItem : ISinglePropertyMappingItem
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

        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, ReadResult mapResult)
        {
            if (!DateTime.TryParseExact(mapResult.StringValue, Formats, Provider, Style, out DateTime result))
            {
                return PropertyMappingResult.Invalid();
            }

            return PropertyMappingResult.Success(result);
        }
    }
}
