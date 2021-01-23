using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper
{
    public class ExcelMappingException : Exception
    {
        /// <summary>
        /// Creates an ExcelMappingException with the default message.
        /// </summary>
        public ExcelMappingException()
        {
        }

        /// <summary>
        /// Creates an ExcelMappingException with the given message.
        /// </summary>
        /// <param name="message">The message of the exception.</param>
        public ExcelMappingException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates an ExcelMappingException with the given message and inner exception.
        /// </summary>
        /// <param name="message">The message of the exception.</param>
        /// <param name="innerException">The inner exception of the exception.</param>
        public ExcelMappingException(string message, Exception innerException) : base(message, innerException)
        {
        }

        /// <summary>
        /// Creates an ExcelMappingException throw trying to map a cell value to a property or field.
        /// </summary>
        /// <param name="message">The base error message of the exception.</param>
        /// <param name="sheet">The sheet that is currently being read.</param>
        /// <param name="rowIndex">The zero-based index of the row in the sheet that is currently being read.</param>
        public ExcelMappingException(string message, ExcelSheet sheet, int rowIndex, int columnIndex) : base(GetMessage(message, sheet, rowIndex, columnIndex))
        {
        }

        /// <summary>
        /// Creates an ExcelMappingException throw trying to map a cell value to a property or field.
        /// </summary>
        /// <param name="message">The base error message of the exception.</param>
        /// <param name="sheet">The sheet that is currently being read.</param>
        /// <param name="rowIndex">The zero-based index of the row in the sheet that is currently being read.</param>
        /// <param name="innerException">The inner exception of the exception.</param>
        public ExcelMappingException(string message, ExcelSheet sheet, int rowIndex, int columnIndex, Exception innerException) : base(GetMessage(message, sheet, rowIndex, columnIndex), innerException)
        {
        }

        private static string GetMessage(string message, ExcelSheet sheet, int rowIndex, int columnIndex)
        {
            string position = string.Empty;
            if (columnIndex != -1)
            {
                if (sheet != null && sheet.HasHeading)
                {
                    if (sheet.Heading == null)
                    {
                        sheet.ReadHeading();
                    }

                    position = $" in column \"{sheet.Heading.GetColumnName(columnIndex)}\"";
                }
                else
                {
                    position = $" in position \"{columnIndex}\"";
                }
            }

            return $"{message}{position} on row {rowIndex} in sheet \"{sheet?.Name}\".";
        }
    }
}
