namespace ExcelMapper;

/// <summary>
/// Converts the value of a cell to a <typeparamref name="T"/>. Set it with
/// <c>WithConverter(converter)</c>, or <see cref="ExcelConverterAttribute"/>, where it replaces the
/// built-in conversion for the member's type.
/// </summary>
/// <typeparam name="T">The type of the member. A converter for a value type <c>T</c> can also be
/// used for members of type <c>T?</c>.</typeparam>
/// <remarks>
/// Converters aren't called for empty cells: those use the member's empty fallback. A converter can
/// be used for many cells at once on different threads, so it mustn't change its own state.
/// </remarks>
public interface ICellConverter<T>
{
    /// <summary>
    /// Converts a cell's value.
    /// </summary>
    /// <param name="value">The cell's value.</param>
    /// <param name="context">Where the cell is, and the member it's mapped to.</param>
    /// <returns>The converted value, or whether the value is invalid or not handled.</returns>
    CellConversionResult<T> Convert(CellValue value, CellContext context);
}
