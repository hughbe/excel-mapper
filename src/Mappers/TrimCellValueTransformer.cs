using System.Reflection;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;
 
public class TrimCellValueTransformer : ICellValueMapper
{
    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
        => previous.Success(previous.StringValue?.Trim());
}
