using System;
using ExcelMapper.Mappings;

namespace ExcelMapper
{
    public class ExcelMappingException : Exception
    {
        public ExcelMappingException() { }

        public ExcelMappingException(string message) : base(message) { }

        public ExcelMappingException(string message, Exception innerException) : base(message, innerException) { }

        public ExcelMappingException(string message, ExcelSheet sheet, int rowIndex, int columnIndex) : base(GetMessage(message, sheet, rowIndex, columnIndex)) { }

        private static string GetMessage(string message, ExcelSheet sheet, int rowIndex, int columnIndex)
        {
            string position;
            if (sheet.HasHeading)
            {
                position = $"\"{sheet.Heading.GetColumnName(columnIndex)}\"";
            }
            else
            {
                position = $"in position \"{columnIndex}\"";
            }


            return $"{message} {position} on row {rowIndex} in sheet \"{sheet?.Name}\".";
        }
    }
}
