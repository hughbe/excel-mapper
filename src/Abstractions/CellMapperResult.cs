using System;
namespace ExcelMapper.Abstractions;

/// <summary>
/// An enumeration describing the result of an operation to map the value
/// of a cell to a property or field.
/// </summary>
public readonly struct CellMapperResult
{
    public object? Value { get; }
    public Exception? Exception { get; }
    public HandleAction Action { get; }

    public bool Succeeded => Action != HandleAction.ErrorAndContinueMapping && Action != HandleAction.IgnoreResultAndContinueMapping;

    internal CellMapperResult(object? value, Exception? exception, HandleAction action)
    {
        Value = value;
        Exception = exception;
        Action = action;
    }

    /// <summary>
    /// The value could be mapped. This mapped value will be used to set the value of the property or
    /// field.
    /// </summary>
    public static CellMapperResult Success(object? value) => new(value, null, HandleAction.UseResultAndStopMapping);

    /// <summary>
    /// The value could be mapped, but prefer the result of mapping items further on in
    /// the mapping pipeline. This can be used for specifiying value mappers that are lower priority.
    /// </summary>
    public static CellMapperResult SuccessIfNoOtherSuccess(object? result) => new(result, null, HandleAction.UseResultAndContinueMapping);

    /// <summary>
    /// The value could not be mapped, but is not invalid. This can be used
    /// for optional value mappers.
    /// </summary>
    public static CellMapperResult Ignore() => new(null, null, HandleAction.IgnoreResultAndContinueMapping);

    /// <summary>
    /// The value was invalid. The InvalidFallback will be invoked if no other value mappers are
    /// successful.
    /// </summary>
    public static CellMapperResult Invalid(Exception exception) => new(null, exception, HandleAction.ErrorAndContinueMapping);

    public enum HandleAction
    {
        /// <summary>
        /// Use the result of mapping. Continue down the pipeline. Handle the success or error when the mapping is
        /// finished only if there are no more subsequent used success or error results.
        /// </summary>
        UseResultAndContinueMapping,

        /// <summary>
        /// Use the result of mapping as an error. Continue down the pipeline. Handle the error when the mapping is
        /// finished only if there are no more subsequent used success or error results.
        /// </summary>
        ErrorAndContinueMapping,

        /// <summary>
        /// Use the result of the mapping. Stop mapping and handle the success or error immediately.
        /// </summary>
        UseResultAndStopMapping,

        /// <summary>
        /// Do not use the result of mapping. Continue down the pipeline.
        /// </summary>
        IgnoreResultAndContinueMapping
    }
}
