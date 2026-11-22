using System.Reflection;

namespace ExcelMapper.Transformers;

/// <summary>
/// Trims text, for <see cref="ExcelTrimStringAttribute"/> and <c>WithTrim()</c>.
/// </summary>
internal sealed class TrimStringCellTransformer : IPipelineTransformer
{
    public ReadCellResult Transform(ReadCellResult cell, ExcelSheet sheet, int rowIndex, MemberInfo? member, bool preserveFormatting)
    {
        // Trimming only changes text. Skip other values, such as dates and numbers, so that they
        // keep their type rather than being converted to a string and parsed again.
        if (!preserveFormatting && cell.GetValue() is not (string or null))
        {
            return cell;
        }

        return new ReadCellResult(cell.ColumnIndex, cell.GetString()?.Trim(), preserveFormatting);
    }
}
