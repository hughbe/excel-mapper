using System;
using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Fallbacks
{
    /// <summary>
    /// A fallback that returns a fixed given value when falling back.
    /// </summary>
    public class FixedValueFallback : IFallbackItem
    {
        /// <summary>
        /// The fixed value returned when falling back.
        /// </summary>
        public object Value { get; }

        /// <summary>
        /// Constructs a fallback that returns a given value when falling back.
        /// </summary>
        /// <param name="value">The fixed value returned when falling back.</param>
        public FixedValueFallback(object value) => Value = value;

        public object PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellValueResult result, Exception exception, MemberInfo member) => Value;
    }
}
