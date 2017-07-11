using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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

        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex, string stringValue)
        {
            if (!DateTime.TryParseExact(stringValue, Formats, Provider, Style, out DateTime result))
            {
                return PropertyMappingResult.Invalid();
            }

            return PropertyMappingResult.Success(result);
        }
    }
}
