using ExcelMapper.Abstractions;

namespace ExcelMapper
{
    /// <summary>
    /// Reads a single cell of an excel sheet and maps the value of the cell to the
    /// type of the property or field.
    /// </summary>
    public class ValuePipeline<T> : ValuePipeline, IValuePipeline<T>
    {
    }
}
