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
        ThrowHelpers.ThrowIfNull(type, nameof(type));

        if (!type.ImplementsInterface(typeof(IConvertible)))
        {
            throw new ArgumentException($"Type \"{type}\" must implement {nameof(IConvertible)} to support {nameof(Convert)}.{nameof(Convert.ChangeType)}.", nameof(type));
        }

        Type = type;
    }

    private static bool IsIntegerType(Type type) => Type.GetTypeCode(type) switch
    {
        TypeCode.SByte or TypeCode.Byte or TypeCode.Int16 or TypeCode.UInt16 or
        TypeCode.Int32 or TypeCode.UInt32 or TypeCode.Int64 or TypeCode.UInt64 => true,
        _ => false
    };

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var value = readResult.GetValue();

        // Excel stores every number as a double. Don't silently round away the fractional part
        // when mapping to an integer type: "2.7" as text can't be parsed as an int, so neither can 2.7.
        if (value is double doubleValue && IsIntegerType(Type) && Math.Floor(doubleValue) != doubleValue
            && !double.IsNaN(doubleValue) && !double.IsInfinity(doubleValue))
        {
            return NotWholeNumberResult(doubleValue, Type);
        }

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

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    private static CellMapperResult NotWholeNumberResult(double value, Type type)
        => CellMapperResult.Invalid(() => new FormatException($"The value {value} is not a whole number, so it can't be converted to {type}."));
}
