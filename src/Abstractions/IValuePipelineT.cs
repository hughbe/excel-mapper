namespace ExcelMapper.Abstractions;

/// <summary>
/// Pipeline for processing cell values through transformers and mappers.
/// </summary>
/// <typeparam name="T">The type of the final mapped value.</typeparam>
public interface IValuePipeline<out T> : IValuePipeline
{
}
