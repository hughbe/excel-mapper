using System;

namespace ExcelMapper.Mappings.Mappers
{
    public delegate PropertyMappingResultType ConvertUsingMappingDelegate(ReadCellValueResult readResult, ref object value);

    /// <summary>
    /// A mapper that tries to map the value of a cell to an object using a given conversion delegate.
    /// </summary>
    public class ConvertUsingMapper : ICellValueMapper
    {
        /// <summary>
        /// Gets the delegate used to map the value of a cell to an object.
        /// </summary>
        public ConvertUsingMappingDelegate Converter { get; }

        /// <summary>
        /// Constructs a mapper that tries to map the value of a cell to an object using a given conversion delegate.
        /// </summary>
        /// <param name="converter">The delegate used to map the value of a cell to an object</param>
        public ConvertUsingMapper(ConvertUsingMappingDelegate converter)
        {
            Converter = converter ?? throw new ArgumentNullException(nameof(converter));
        }

        public PropertyMappingResultType GetProperty(ReadCellValueResult readResult, ref object value)
        {
            return Converter(readResult, ref value);
        }
    }
}
