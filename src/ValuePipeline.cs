using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper;

/// <summary>
/// Reads a single cell of an excel sheet and maps the value of the cell to the
/// type of the property or field.
/// </summary>
public class ValuePipeline : IValuePipeline
{
    /// <summary>
    /// Gets the list of objects that take the initial string value read from a cell and
    /// modifies the string value. This is useful for things like trimming the string value
    /// before mapping it.
    /// </summary>
    public IList<ICellTransformer> Transformers { get; } = new NonNullCollection<ICellTransformer>();

    /// <summary>
    /// Gets the pipeline of items that take the initial string value read from a cell and
    /// converts the string value into the type of the property or field. The items form
    /// a pipeline: if a mapper fails to parse or map the cell value, the next item is used.
    /// </summary>
    public IList<ICellMapper> Mappers { get; } = new NonNullCollection<ICellMapper>();

    /// <summary>
    /// Gets or sets an object that handles mapping a cell value to a property or field if the value of the
    /// cell is empty. For example, you can provide a fixed value to return if the value of the cell
    /// is empty.
    /// </summary>
    public IFallbackItem? EmptyFallback { get; set; }

    /// <summary>
    /// Gets or sets an object that handles mapping a cell value to a property or field if all items
    /// in the mapper pipeline failed to map the value to the property or field. For example, you can
    /// provide a fixed value to return if the value of the cell is invalid.
    /// </summary>
    public IFallbackItem? InvalidFallback { get; set; }

    /// <summary>
    /// Processes a cell value through the complete transformation and mapping pipeline to produce a final property value.
    /// This method orchestrates the entire value mapping process: transforming the raw cell value, running it through
    /// the mapper pipeline, and applying fallback strategies when needed.
    /// </summary>
    /// <param name="pipeline">The value pipeline containing transformers, mappers, and fallback items to use.</param>
    /// <param name="sheet">The Excel sheet being read, used for context in transformers and error messages.</param>
    /// <param name="rowIndex">The zero-based index of the row being processed, used for error reporting.</param>
    /// <param name="readResult">The result of reading a cell, containing the raw value and column information.</param>
    /// <param name="preserveFormatting">Whether to preserve Excel formatting when reading string values.</param>
    /// <param name="member">The property or field being mapped to, used for error messages. Can be null.</param>
    /// <returns>
    /// The final mapped value after applying transformers, mappers, and fallbacks. 
    /// May return null if empty fallback returns null, or throw an exception if fallbacks are configured to throw.
    /// </returns>
    /// <remarks>
    /// <para>Processing order:</para>
    /// <list type="number">
    /// <item><description>Apply all cell value transformers (e.g., trimming)</description></item>
    /// <item><description>Check if value is empty and apply empty fallback if configured</description></item>
    /// <item><description>Run value through mapper pipeline until one succeeds or all fail</description></item>
    /// <item><description>If all mappers fail, apply invalid fallback if configured</description></item>
    /// <item><description>Return the final mapped value or throw if no fallback handled the failure</description></item>
    /// </list>
    /// </remarks>
    internal static object? GetPropertyValue(
        IValuePipeline pipeline,
        ExcelSheet sheet,
        int rowIndex,
        ReadCellResult readResult,
        bool preserveFormatting,
        MemberInfo? member)
    {
        foreach (var transformer in pipeline.Transformers)
        {
            readResult = new ReadCellResult(readResult.ColumnIndex, transformer.TransformStringValue(sheet, rowIndex, readResult), preserveFormatting);
        }

        if (readResult.IsEmpty() && pipeline.EmptyFallback != null)
        {
            return pipeline.EmptyFallback.PerformFallback(sheet, rowIndex, readResult, null, member);
        }

        CellMapperResult? finalResult = null;
        foreach (var mapper in pipeline.Mappers)
        {
            var result = mapper.Map(readResult);
            if (result.Action != CellMapperResult.HandleAction.IgnoreResultAndContinueMapping)
            {
                finalResult = result;
            }

            if (result.Action == CellMapperResult.HandleAction.UseResultAndStopMapping)
            {
                break;
            }
        }


        if ((finalResult == null || !finalResult.Value.Succeeded) && pipeline.InvalidFallback != null)
        {
            return pipeline.InvalidFallback.PerformFallback(sheet, rowIndex, readResult, finalResult?.Exception, member);
        }

        return finalResult?.Value;
    }
}
