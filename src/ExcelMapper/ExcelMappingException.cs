using System;

namespace ExcelMapper
{
    public class ExcelMappingException : Exception
    {
        public ExcelMappingException() { }

        public ExcelMappingException(string message) : base(message) { }

        public ExcelMappingException(string message, Exception innerException) : base(message, innerException) { }

        public ExcelMappingException(string message, string position, ExcelSheet sheet, ExcelRow row) : base($"{message} for {position} on row {row.Index} in sheet \"{sheet?.Name}\"")
        {
        }
    }
}
