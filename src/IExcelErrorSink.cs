namespace ExcelMapper;

/// <summary>
/// Receives the problems found while rows are read with
/// <see cref="ExcelSheet.ReadRows{T}(IExcelErrorSink)"/>, as they're found. Unlike
/// <see cref="ExcelSheet.ReadRowResults{T}()"/>, no exception is created for an empty or invalid
/// cell unless the sink asks for one, so checking a large file with many invalid cells doesn't
/// allocate for each of them.
/// </summary>
/// <example>
/// <code>
/// sealed class ErrorCounter : IExcelErrorSink
/// {
///     public int Count;
///     public void Add(in ExcelMappingError error) => Count++;
/// }
/// </code>
/// </example>
public interface IExcelErrorSink
{
    /// <summary>
    /// Receives a problem with a row. The error is only valid during the call: copy what you need
    /// from it, or call <see cref="ExcelMappingError.ToException"/> to keep it.
    /// </summary>
    /// <param name="error">The problem.</param>
    void Add(in ExcelMappingError error);
}
