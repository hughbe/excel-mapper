namespace ExcelMapper;

/// <summary>
/// The result of mapping a row with <see cref="ExcelSheet.ReadRowResults{T}"/>: either the mapped
/// value, or every error that prevented the row from being mapped.
/// </summary>
/// <typeparam name="T">The type of the object the row is mapped to.</typeparam>
public sealed class ExcelRowResult<T>
{
    internal ExcelRowResult(int rowIndex, T? value, IReadOnlyList<ExcelMappingException> errors)
    {
        RowIndex = rowIndex;
        Value = value;
        Errors = errors;
    }

    /// <summary>
    /// Gets the zero-based index of the row in the sheet.
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// Gets the mapped value, or the default value of <typeparamref name="T"/> if the row is invalid.
    /// </summary>
    public T? Value { get; }

    /// <summary>
    /// Gets the errors that prevented the row from being mapped. Empty if the row is valid. Each
    /// error's <see cref="ExcelMappingException.CellAddress"/> gives the cell, where known.
    /// </summary>
    public IReadOnlyList<ExcelMappingException> Errors { get; }

    /// <summary>
    /// Gets whether the row was mapped without errors.
    /// </summary>
    public bool IsValid => Errors.Count == 0;
}
