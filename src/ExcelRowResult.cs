namespace ExcelMapper;

/// <summary>
/// The result of mapping a row with <see cref="ExcelSheet.ReadRowResults{T}()"/>: either the mapped
/// value, or every error that prevented the row from being mapped.
/// </summary>
/// <typeparam name="T">The type of the object the row is mapped to.</typeparam>
public readonly struct ExcelRowResult<T>
{
    private readonly IReadOnlyList<ExcelMappingException>? _errors;

    internal ExcelRowResult(int rowIndex, T? value, IReadOnlyList<ExcelMappingException>? errors)
    {
        RowIndex = rowIndex;
        Value = value;
        _errors = errors;
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
    public IReadOnlyList<ExcelMappingException> Errors => _errors ?? Array.Empty<ExcelMappingException>();

    /// <summary>
    /// Gets whether the row was mapped without errors.
    /// </summary>
    public bool IsValid => _errors is null || _errors.Count == 0;
}
