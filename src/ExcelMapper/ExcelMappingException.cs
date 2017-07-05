using System;
using ExcelMapper.Pipeline;

namespace ExcelMapper
{
    public class ExcelMappingException : Exception
    {
        public ExcelMappingException() { }

        public ExcelMappingException(string message) : base(message) { }

        public ExcelMappingException(string message, Exception innerException) : base(message, innerException) { }

        public ExcelMappingException(string message, PipelineContext context) : base(GetMessage(message, context)) { }

        private static string GetMessage(string message, PipelineContext context)
        {
            string position;
            if (context.Sheet.HasHeading)
            {
                position = $"\"{context.Sheet.Heading.GetColumnName(context.ColumnIndex)}\"";
            }
            else
            {
                position = $"in position \"{context.ColumnIndex}\"";
            }


            return $"{message} {position} on row {context.RowIndex} in sheet \"{context.Sheet?.Name}\".";
        }
    }
}
