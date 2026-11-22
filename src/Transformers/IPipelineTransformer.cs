using System.Reflection;

namespace ExcelMapper.Transformers;

/// <summary>
/// Changes a cell before it's converted, as the first stage of a <see cref="ValuePipeline"/>. The
/// built-in trimming and the adapter for <see cref="ExcelMapper.ICellTransformer"/> implement it.
/// </summary>
internal interface IPipelineTransformer
{
    /// <param name="cell">The cell, after any earlier transformers.</param>
    /// <param name="sheet">The sheet being read.</param>
    /// <param name="rowIndex">The index of the row being read.</param>
    /// <param name="member">The member being mapped, if any.</param>
    /// <param name="preserveFormatting">Whether the member reads the cell's formatted text.</param>
    ReadCellResult Transform(ReadCellResult cell, ExcelSheet sheet, int rowIndex, MemberInfo? member, bool preserveFormatting);
}
