namespace ExcelMapper.Fallbacks;

/// <summary>
/// The default invalid fallback, and the empty fallback of non-nullable value types: reports the
/// cell as an error.
/// </summary>
internal sealed class ThrowFallback<T> : ICellFallback<T>
{
    public static ThrowFallback<T> Instance { get; } = new();

    // What to do about an empty cell. A type that can be null can already be empty.
    private static readonly string s_emptyHint = default(T) is null
        ? " Give it an empty fallback, such as [ExcelDefaultValue]."
        : $" Make it nullable ({TypeNames.Get(typeof(T))}?), or give it an empty fallback, such as [ExcelDefaultValue].";

    public T Fallback(CellValue value, CellContext context, Exception? exception)
    {
        var mappingException = CreateException(value, context, exception);

        // When errors are being collected (ReadRowResults), record the error rather than throwing
        // it: throwing and catching an exception for every invalid cell is slow. The member isn't
        // set, as the row has an error.
        if (context.Errors is { } errors)
        {
            errors.Add(mappingException);
            return default!;
        }

        throw mappingException;
    }

    // Says what was wrong with the cell: that it's empty, or why its value isn't valid.
    private static ExcelMappingException CreateException(CellValue value, CellContext context, Exception? exception)
    {
        var member = context.Member;
        var type = TypeNames.Get(member?.MemberType() ?? typeof(T));
        string message;
        string? hint;
        if (value.IsEmpty)
        {
            message = member is null
                ? $"The cell is empty, but type {type} can't be empty"
                : $"The cell is empty, but member \"{member.Name}\" of type {type} can't be empty";
            hint = s_emptyHint;
        }
        else
        {
            message = member is null
                ? $"Cannot assign \"{value.Text}\" to type {type}"
                : $"Cannot assign \"{value.Text}\" to member \"{member.Name}\" of type {type}";

            // Out-of-range numbers say why, as the text looks like a valid number.
            hint = exception is OverflowException ? " " + exception.Message : null;
        }

        return new ExcelMappingException(message, context.View, context.RowIndex, context.ColumnIndex, exception, hint)
        {
            Member = member
        };
    }
}
