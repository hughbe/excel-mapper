namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to an IConvertible object using Convert.ChangeType.
/// Numbers and the other built-in types have their own mappers, so this is used for other IConvertible types.
/// </summary>
/// <typeparam name="T">The member's type: the IConvertible type, or a nullable version of it.</typeparam>
internal sealed class ChangeTypeMapper<T> : ICellMapper<T>
{
    private static readonly Type s_type = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);

    public CellConversion Convert(ReadCellResult cell, in CellContext context, out T value, ref object? error)
    {
        try
        {
            value = (T)System.Convert.ChangeType(cell.GetValue(), s_type, cell.Culture)!;
            return CellConversion.Success;
        }
        catch
        {
            value = default!;
            return CellConversion.Invalid;
        }
    }

    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error)
        => MapperUtilities.InvalidFromParse(() => System.Convert.ChangeType(cell.GetValue(), s_type, cell.Culture));
}
