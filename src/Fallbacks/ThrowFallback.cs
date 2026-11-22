namespace ExcelMapper.Fallbacks;

/// <summary>
/// The default invalid fallback, and the empty fallback of non-nullable value types: reports the
/// cell as an error.
/// </summary>
internal sealed class ThrowFallback<T> : ICellFallback<T>
{
    public static ThrowFallback<T> Instance { get; } = new();

    public T Fallback(CellValue value, CellContext context, Exception? exception)
    {
        var member = context.Member;
        var sheet = context.View;
        var text = value.Text;
        var mappingException = member == null
            ? new ExcelMappingException($"Cannot assign \"{text}\"", sheet, context.RowIndex, context.ColumnIndex, exception)
            : new ExcelMappingException($"Cannot assign \"{text}\" to member \"{member.Name}\" of type \"{member.MemberType()}\"", sheet, context.RowIndex, context.ColumnIndex, exception)
            {
                Member = member
            };

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
}
