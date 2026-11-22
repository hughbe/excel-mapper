using System.Reflection;

namespace ExcelMapper.Transformers;

/// <summary>
/// The pipeline transformer for an <see cref="ExcelMapper.ICellTransformer"/>.
/// </summary>
internal sealed class CellTransformerAdapter(ExcelMapper.ICellTransformer transformer) : IPipelineTransformer
{
    public ExcelMapper.ICellTransformer Transformer { get; } = transformer;

    public ReadCellResult Transform(ReadCellResult cell, ExcelSheet sheet, int rowIndex, MemberInfo? member, bool preserveFormatting)
        => Transformer.Transform(new CellValue(cell), new CellContext(sheet, rowIndex, cell.ColumnIndex, member)).ToReadCellResult(cell.ColumnIndex);
}
