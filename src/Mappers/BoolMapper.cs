namespace ExcelMapper.Mappers;

internal sealed class BoolMapper : ICellMapper<bool>, ICellMapper<bool?>
{
    CellConversion ICellMapper<bool>.Convert(ReadCellResult cell, in CellContext context, out bool value, ref object? error)
    {
        if (cell.GetValue() is bool boolValue)
        {
            value = boolValue;
            return CellConversion.Success;
        }

        // Excel transforms bool values such as "true" or "false" to "1" or "0", and spreadsheets
        // often use "Yes" and "No".
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

    CellConversion ICellMapper<bool?>.Convert(ReadCellResult cell, in CellContext context, out bool? value, ref object? error)
    {
        var conversion = ((ICellMapper<bool>)this).Convert(cell, in context, out var result, ref error);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error)
    {
        var stringValue = cell.GetString();
        return MapperUtilities.InvalidFormat(stringValue, "a valid Boolean", () => bool.Parse(stringValue!));
    }
}
