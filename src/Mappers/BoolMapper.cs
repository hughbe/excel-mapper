namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a bool.
/// </summary>
public class BoolMapper : ICellMapper
{
    private static readonly object s_boxedTrue = true;
    private static readonly object s_boxedFalse = false;

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        if (readResult.GetValue() is bool boolValue)
        {
            return CellMapperResult.Success(boolValue);
        }

        var stringValue = readResult.GetString();

        // Excel transforms bool values such as "true" or "false" to "1" or "0".
        if (stringValue == "1")
        {
            return CellMapperResult.Success(s_boxedTrue);
        }
        if (stringValue == "0")
        {
            return CellMapperResult.Success(s_boxedFalse);
        }

        if (bool.TryParse(stringValue, out var result))
        {
            return CellMapperResult.Success(result ? s_boxedTrue : s_boxedFalse);
        }

        // Only create the exception (by parsing again) if it is needed. If stringValue is null this
        // is an ArgumentNullException.
        return MapperUtilities.InvalidFromParse(() => bool.Parse(stringValue!));
    }
}
