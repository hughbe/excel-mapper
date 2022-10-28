using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

public class InvalidCellMapper : ICellValueMapper
{
    public IInvalidCellFallback Fallback { get; }

    public InvalidCellMapper(IInvalidCellFallback fallback)
    {
        Fallback = fallback ?? throw new ArgumentNullException(nameof(fallback));
    }

    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        if (!previous.Succeeded)
        {
            return previous.Success(Fallback.PerformFallback(cell, previous.Value, previous.Exception, member));
        }

        return previous;
    }
}
