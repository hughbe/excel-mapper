using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Readers;

namespace ExcelMapper.Abstractions;

/// <summary>
/// A map that reads several cells (a collection or dictionary) and can be bound to a member: its
/// cells reader is resolved once for the sheet.
/// </summary>
internal interface ICellsMap : IColumnsMap
{
    /// <summary>
    /// Gets the reader of the map's cells in a sheet, or null if the sheet doesn't have them.
    /// </summary>
    ICellsReader? GetCellsReader(SheetView sheet);

    /// <summary>
    /// Reads the collection or dictionary from the current row with a reader from
    /// <see cref="GetCellsReader"/>, as <see cref="IMap.TryGetValue"/> does.
    /// </summary>
    bool TryGetValue(ICellsReader? cellsReader, SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, out object? value);
}
