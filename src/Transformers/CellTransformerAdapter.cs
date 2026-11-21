using System.Reflection;

namespace ExcelMapper.Transformers;

/// <summary>
/// The pipeline transformer for an <see cref="ExcelMapper.ICellTransformer"/>. Pipelines call
/// <see cref="Transform(ReadCellResult, ExcelSheet, int, MemberInfo?)"/>, which keeps the value's
/// type; <see cref="TransformStringValue"/> is for callers of the older interface.
/// </summary>
internal sealed class CellTransformerAdapter(ExcelMapper.ICellTransformer transformer) : Abstractions.ICellTransformer
{
    public ExcelMapper.ICellTransformer Transformer { get; } = transformer;

    public ReadCellResult Transform(ReadCellResult readResult, ExcelSheet sheet, int rowIndex, MemberInfo? member)
        => Transformer.Transform(new CellValue(readResult), new CellContext(sheet, rowIndex, readResult.ColumnIndex, member)).ToReadCellResult(readResult.ColumnIndex);

    public string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
        => Transform(readResult, sheet, rowIndex, null).GetString();
}
