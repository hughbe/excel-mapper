namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to an IConvertible object using Convert.ChangeType.
/// Numbers and the other built-in types have their own mappers, so this is used for other IConvertible types.
/// </summary>
internal class ChangeTypeMapper : ICellMapper
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
        Type = type;
    }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
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
