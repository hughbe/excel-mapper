using System.Linq;
using System.Reflection;

namespace ExcelMapper.Readers;

/// <summary>
/// A column with a name, such as <c>[ExcelColumnName("Name")]</c> or the member's own name.
/// </summary>
internal sealed class ColumnNameColumns : Columns
{
    public ColumnNameColumns(string columnName, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        ColumnUtilities.ValidateColumnName(columnName);
        EnumUtilities.ValidateIsDefined(comparison);
        Name = columnName;
        Comparison = comparison;
    }

    public string Name { get; }

    public StringComparison Comparison { get; }

    public override int FindColumn(SheetView sheet)
        => ColumnUtilities.TryGetColumnIndex(sheet, Name, Comparison, out var column) ? column : -1;

    public override bool LooksUpNames => true;

    public override string ColumnName => Name;

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, member, $"column \"{Name}\"", ExcelMappingException.GetSimilarColumnHint(sheet, Name));

    public override string Describe(SheetView sheet)
    {
        var similar = ExcelMappingException.GetSimilarColumnHint(sheet, Name) is null ? null : SimilarColumn(sheet, Name);
        return similar is null ? $"column \"{Name}\"" : $"column \"{Name}\", did you mean \"{similar}\"?";
    }

    private static string SimilarColumn(SheetView sheet, string columnName)
    {
        var normalizedName = ColumnUtilities.RemoveSeparators(columnName);
        return sheet.Heading!.ColumnNames.First(c => string.Equals(ColumnUtilities.RemoveSeparators(c), normalizedName, StringComparison.OrdinalIgnoreCase));
    }
}
