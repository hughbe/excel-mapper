using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

public class EmptyCellMapper : ICellValueMapper
{
    public IEmptyCellFallback Fallback { get; }

    public EmptyCellMapper(IEmptyCellFallback fallback)
    {
        Fallback = fallback ?? throw new ArgumentNullException(nameof(fallback));
    }

    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        if (previous.Value == null || (previous.Value is string stringValue && string.IsNullOrEmpty(stringValue)))
        {
            return previous.Success(Fallback.PerformFallback(cell, member));
        }

        return previous;
    }
}
