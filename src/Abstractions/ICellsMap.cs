using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Abstractions;

/// <summary>
/// A map that reads several cells (a collection or dictionary) and can be bound to a member: its
/// cells reader is resolved once for the sheet.
/// </summary>
internal interface ICellsMap : IColumnsMap
{
    ICellsReader? GetCellsReader(SheetView sheet);

    bool TryGetValue(ICellsReader? cellsReader, SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, out object? value);
}
