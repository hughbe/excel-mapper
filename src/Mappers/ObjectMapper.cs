namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that returns the value of a cell as Excel stores it: for example a <see cref="double"/>
/// for numbers, a <see cref="DateTime"/> for dates, a <see cref="bool"/> for booleans and a
/// <see cref="string"/> for text. If formatting is preserved, the displayed text is returned.
/// This is used for <see cref="object"/> and <see cref="IConvertible"/> members.
/// </summary>
internal class ObjectMapper : ICellMapper
{
    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
        => CellMapperResult.SuccessIfNoOtherSuccess(readResult.PreserveFormatting ? readResult.GetString() : readResult.GetValue());
}
