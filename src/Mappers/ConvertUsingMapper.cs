using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers
{
    public delegate CellValueMapperResult ConvertUsingMapperDelegate(ReadCellValueResult readResult);

    /// <summary>
    /// A mapper that tries to map the value of a cell to an object using a given conversion delegate.
    /// </summary>
    public class ConvertUsingMapper : ICellValueMapper
    {
        /// <summary>
        /// Gets the delegate used to map the value of a cell to an object.
        /// </summary>
        public ConvertUsingMapperDelegate Converter { get; }

        /// <summary>
        /// Constructs a mapper that tries to map the value of a cell to an object using a given conversion delegate.
        /// </summary>
        /// <param name="converter">The delegate used to map the value of a cell to an object</param>
        public ConvertUsingMapper(ConvertUsingMapperDelegate converter)
        {
            Converter = converter ?? throw new ArgumentNullException(nameof(converter));
        }

        public CellValueMapperResult MapCellValue(ReadCellValueResult readResult)
        {
            try
            {
                return Converter(readResult);
            }
            catch (Exception exception)
            {
                return CellValueMapperResult.Invalid(exception);
            }
        }
    }
}
