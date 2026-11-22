namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that returns the value of a cell as Excel stores it: for example a <see cref="double"/>
/// for numbers, a <see cref="DateTime"/> for dates, a <see cref="bool"/> for booleans and a
/// <see cref="string"/> for text. With PreserveFormatting, it returns the formatted text instead.
/// Used for <see cref="object"/> and <see cref="IConvertible"/> members.
/// </summary>
internal sealed class ObjectMapper<T> : ICellMapper<T>
{
    public CellConversion Convert(ReadCellResult cell, out T value)
    {
        value = (T)(cell.PreserveFormatting ? cell.GetString() : cell.GetValue())!;
        return CellConversion.SuccessUnlessReplaced;
    }

    // Every value is valid.
    public Exception CreateException(ReadCellResult cell) => new FormatException("The value couldn't be read.");
}
