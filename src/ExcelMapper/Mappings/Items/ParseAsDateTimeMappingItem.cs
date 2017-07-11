using System;
using System.Globalization;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    internal class ParseAsDateTimeMappingItem : ISinglePropertyMappingItem
    {
        /// <summary>
        /// Defaults to "G" - the default Excel format.
        /// </summary>
        public string[] Formats { get; internal set; } = new string[] { "G" };

        public IFormatProvider Provider { get; internal set; }

        public DateTimeStyles Style { get; internal set; }

        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MapResult mapResult)
        {
            if (!DateTime.TryParseExact(mapResult.StringValue, Formats, Provider, Style, out DateTime result))
            {
                return PropertyMappingResult.Invalid();
            }

            return PropertyMappingResult.Success(result);
        }
    }
}
