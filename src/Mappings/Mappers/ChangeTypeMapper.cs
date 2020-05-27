using System;
using ExcelMapper.Utilities;

namespace ExcelMapper.Mappings.Mappers
{
    /// <summary>
    /// A mapper that tries to map the value of a cell to an IConvertible object using Convert.ChangeType.
    /// </summary>
    public class ChangeTypeMapper : ICellValueMapper
    {
        /// <summary>
        /// Gets the type of the IConvertible object to map the value of a cell to.
        /// </summary>
        public Type Type { get; }

        /// <summary>
        /// Constructs a mapper that tries to map the value of a cell to an IConvertible object using
        /// Convert.ChangeType.
        /// </summary>
        /// <param name="type">The type of the IConvertible object to map the value of a cell to.</param>
        public ChangeTypeMapper(Type type)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            if (!type.ImplementsInterface(typeof(IConvertible)))
            {
                throw new ArgumentException($"Type \"{type}\" must implement IConvertible to support Convert.ChangeType.", nameof(type));
            }

            Type = type;
        }

        public PropertyMapperResultType MapCellValue(ReadCellValueResult readResult, ref object value)
        {
            try
            {
                value = Convert.ChangeType(readResult.StringValue, Type);
                return PropertyMapperResultType.Success;
            }
            catch
            {
                return PropertyMapperResultType.Invalid;
            }
        }
    }
}
