namespace ExcelMapper.Utilities;

/// <summary>
/// Creates the exception for an empty or invalid cell that its member can't take: thrown by
/// <see cref="Fallbacks.ThrowFallback{T}"/>, collected by <see cref="ExcelSheet.ReadRowResults{T}()"/>,
/// or created when an <see cref="IExcelErrorSink"/> asks (<see cref="ExcelMappingError.ToException"/>).
/// </summary>
internal static class CellErrors
{
    /// <summary>
    /// Gets what to do about an empty cell of <paramref name="type"/>: a type that can be null can
    /// already be empty, so it needs an empty fallback.
    /// </summary>
    public static string GetEmptyHint(Type type)
        => !type.IsValueType || Nullable.GetUnderlyingType(type) is not null
            ? " Give it an empty fallback, such as [ExcelDefaultValue]."
            : $" Make it nullable ({TypeNames.Get(type)}?), or give it an empty fallback, such as [ExcelDefaultValue].";

    /// <summary>
    /// Says what was wrong with the cell: that it's empty, or why its value isn't valid.
    /// </summary>
    /// <param name="value">The cell.</param>
    /// <param name="context">The cell's context.</param>
    /// <param name="exception">Why the value is invalid, if known.</param>
    /// <param name="valueType">The type the cell was converted to.</param>
    /// <param name="emptyHint">What to do about an empty cell (see <see cref="GetEmptyHint"/>).</param>
    public static ExcelMappingException Create(CellValue value, in CellContext context, Exception? exception, Type valueType, string emptyHint)
    {
        var member = context.Member;
        var type = TypeNames.Get(member?.MemberType() ?? valueType);
        string message;
        string? hint;
        if (value.IsEmpty)
        {
            message = member is null
                ? $"The cell is empty, but type {type} can't be empty"
                : $"The cell is empty, but member \"{member.Name}\" of type {type} can't be empty";
            hint = emptyHint;
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
