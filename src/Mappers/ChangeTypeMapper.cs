using System;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to an IConvertible object using Convert.ChangeType.
/// </summary>
public class ChangeTypeMapper : ICellMapper
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
        ArgumentNullException.ThrowIfNull(type);

        if (!type.ImplementsInterface(typeof(IConvertible)))
        {
            throw new ArgumentException($"Type \"{type}\" must implement {nameof(IConvertible)} to support {nameof(Convert)}.{nameof(Convert.ChangeType)}.", nameof(type));
        }

        Type = type;
    }

    public CellMapperResult MapCellValue(ReadCellResult readResult)
    {
        var value = readResult.GetValue();
        try
        {
            var result = Convert.ChangeType(value, Type);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}
