using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Utilities;
 
/// <summary>
/// Reads a single cell of an excel sheet and maps the value of the cell to the
/// type of the property or field.
/// </summary>
internal class ValuePipeline
{
    private static Exception s_couldNotMapException = new ExcelMappingException("Could not map successfully.");

    internal static object GetPropertyValue(
        ExcelCell cell,
        object cellValue,
        MemberInfo member,
        IEnumerable<ICellValueMapper> mappers
    )
    {
        var previousResult = new CellValueMapperResult(cellValue, s_couldNotMapException, CellValueMapperResult.HandleAction.UseResultAndContinueMapping);
        foreach (ICellValueMapper mapper in mappers)
        {
            CellValueMapperResult result = mapper.MapCell(cell, previousResult, member);
            if (result.Action != CellValueMapperResult.HandleAction.IgnoreResultAndContinueMapping)
            {
                previousResult = result;
            }

            if (result.Action == CellValueMapperResult.HandleAction.UseResultAndStopMapping)
            {
                break;
            }
        }

        return previousResult.Value;
    }
}
