namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a bool.
/// </summary>
public class BoolMapper : ICellMapper, ICellConverter<bool>, ICellConverter<bool?>
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

        // Spreadsheets often use "Yes" and "No".
        if (string.Equals(stringValue, "yes", StringComparison.OrdinalIgnoreCase))
        {
            return CellMapperResult.Success(s_boxedTrue);
        }
        if (string.Equals(stringValue, "no", StringComparison.OrdinalIgnoreCase))
        {
            return CellMapperResult.Success(s_boxedFalse);
        }

        if (bool.TryParse(stringValue, out var result))
        {
            return CellMapperResult.Success(result ? s_boxedTrue : s_boxedFalse);
        }

        // Only create the exception (by parsing again) if it is needed. If stringValue is null this
        // is an ArgumentNullException.
        return InvalidResult(stringValue);
    }

    CellConversion ICellConverter<bool>.Convert(ReadCellResult cell, out bool value)
    {
        if (cell.GetValue() is bool boolValue)
        {
            value = boolValue;
            return CellConversion.Success;
        }

        var stringValue = cell.GetString();
        if (stringValue == "1" || string.Equals(stringValue, "yes", StringComparison.OrdinalIgnoreCase))
        {
            value = true;
            return CellConversion.Success;
        }

        if (stringValue == "0" || string.Equals(stringValue, "no", StringComparison.OrdinalIgnoreCase))
        {
            value = false;
            return CellConversion.Success;
        }

        return bool.TryParse(stringValue, out value) ? CellConversion.Success : CellConversion.Invalid;
    }

    CellConversion ICellConverter<bool?>.Convert(ReadCellResult cell, out bool? value)
    {
        var conversion = ((ICellConverter<bool>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    private static CellMapperResult InvalidResult(string? stringValue)
        => MapperUtilities.InvalidFromParse(() => bool.Parse(stringValue!));
}
