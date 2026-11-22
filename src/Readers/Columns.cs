using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelMapper.Readers;

/// <summary>
/// The columns a member reads, chosen by name, position, a pattern or an <see cref="IColumnSelector"/>.
/// They're found once for each sheet (see <see cref="SheetView.FindColumn"/>): a member with a
/// single value reads the first column found, and a collection or dictionary reads all of them.
/// </summary>
/// <remarks>
/// Each kind also describes itself, for the error when its columns are missing and for
/// <see cref="ExcelSheet.ValidateHeading{T}"/>.
/// </remarks>
internal abstract class Columns
{
    /// <summary>
    /// Finds the column a single value is read from.
    /// </summary>
    /// <returns>The column's position in the worksheet, or -1 if the sheet doesn't have it.</returns>
    public abstract int FindColumn(SheetView sheet);

    /// <summary>
    /// Finds the columns a collection or dictionary is read from.
    /// </summary>
    /// <returns>The columns' positions in the worksheet, in the order to read them, or null if
    /// the sheet doesn't have them.</returns>
    public virtual int[]? FindColumns(SheetView sheet)
    {
        var column = FindColumn(sheet);
        return column == -1 ? null : [column];
    }

    /// <summary>
    /// Whether the columns are looked up by name, so that they identify a heading row (see
    /// <see cref="ExcelSheet.FindHeading{T}()"/>).
    /// </summary>
    public virtual bool LooksUpNames => false;

    /// <summary>
    /// The column's name, for <see cref="ExcelMappingException.ColumnName"/>, if there's one.
    /// </summary>
    public virtual string? ColumnName => null;

    /// <summary>
    /// Creates the error for a member whose columns the sheet doesn't have.
    /// </summary>
    public virtual ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, rowIndex, member, null);

    /// <summary>
    /// Describes the columns for <see cref="ExcelSheet.ValidateHeading{T}"/>, such as
    /// <c>column "Name"</c>.
    /// </summary>
    public virtual string Describe(SheetView sheet) => "a matching column";

    private protected static ExcelMappingException Missing(SheetView sheet, int rowIndex, MemberInfo? member, string? columns)
        => new($"Could not read value for member \"{member?.Name}\"{(columns is null ? string.Empty : " for " + columns)}", sheet, rowIndex, -1);

    private protected static ExcelMappingException NoHeading(SheetView sheet)
        => new($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");

    private protected static string DescribeNames(IReadOnlyList<string> names)
        => names.Count == 1 ? $"column \"{names[0]}\"" : $"columns {string.Join(", ", names.Select(n => $"\"{n}\""))}";
}

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
        => new($"Could not read value for member \"{member?.Name}\" for column \"{Name}\"", sheet, rowIndex, -1, null, ExcelMappingException.GetSimilarColumnHint(sheet, Name));

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

/// <summary>
/// The first of several columns that the sheet has, from several <see cref="ExcelColumnNameAttribute"/>s.
/// Missing columns are described by the first.
/// </summary>
internal sealed class FirstColumnNameColumns(IReadOnlyList<ColumnNameColumns> columns) : Columns
{
    public IReadOnlyList<ColumnNameColumns> Columns { get; } = columns;

    public override int FindColumn(SheetView sheet)
    {
        for (int i = 0; i < Columns.Count; i++)
        {
            if (Columns[i].FindColumn(sheet) is var column and not -1)
            {
                return column;
            }
        }

        return -1;
    }

    public override bool LooksUpNames => true;

    public override string ColumnName => Columns[0].Name;

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Columns[0].CreateMissingException(sheet, rowIndex, member);

    public override string Describe(SheetView sheet) => Columns[0].Describe(sheet);
}

/// <summary>
/// Columns with names: a single value reads the first the sheet has, and a collection or
/// dictionary reads all of them.
/// </summary>
internal sealed class ColumnNamesColumns : Columns
{
    public ColumnNamesColumns(IReadOnlyList<string> columnNames, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        ColumnUtilities.ValidateColumnNames(columnNames);
        EnumUtilities.ValidateIsDefined(comparison);
        Names = columnNames;
        Comparison = comparison;
    }

    public IReadOnlyList<string> Names { get; }

    public StringComparison Comparison { get; }

    public override int FindColumn(SheetView sheet)
    {
        foreach (var name in Names)
        {
            if (ColumnUtilities.TryGetColumnIndex(sheet, name, Comparison, out var column))
            {
                return column;
            }
        }

        return -1;
    }

    public override int[]? FindColumns(SheetView sheet)
    {
        if (sheet.Heading == null)
        {
            throw NoHeading(sheet);
        }

        var columns = new int[Names.Count];
        for (int i = 0; i < columns.Length; i++)
        {
            if (!ColumnUtilities.TryGetColumnIndex(sheet, Names[i], Comparison, out columns[i]))
            {
                return null;
            }
        }

        return columns;
    }

    public override bool LooksUpNames => true;

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, rowIndex, member, DescribeNames(Names));

    public override string Describe(SheetView sheet) => $"one of columns {string.Join(", ", Names.Select(n => $"\"{n}\""))}";
}

/// <summary>
/// A column at a position, such as <c>[ExcelColumnIndex(2)]</c>. In a region, positions count from
/// the region's first column.
/// </summary>
internal sealed class ColumnIndexColumns : Columns
{
    public ColumnIndexColumns(int columnIndex)
    {
        ThrowHelpers.ThrowIfNegative(columnIndex, nameof(columnIndex));
        Index = columnIndex;
    }

    public int Index { get; }

    public override int FindColumn(SheetView sheet) => sheet.ToSheetColumn(Index);

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, rowIndex, member, $"column at index {Index}");

    public override string Describe(SheetView sheet)
        => $"column {ExcelRange.GetColumnLetters(sheet.ColumnOffset + Index)} (index {Index})";
}

/// <summary>
/// Columns at positions: a single value reads the first the sheet has, and a collection or
/// dictionary reads all of them.
/// </summary>
internal sealed class ColumnIndicesColumns : Columns
{
    public ColumnIndicesColumns(IReadOnlyList<int> columnIndices)
    {
        ColumnUtilities.ValidateColumnIndices(columnIndices);
        Indices = columnIndices;
    }

    public IReadOnlyList<int> Indices { get; }

    public override int FindColumn(SheetView sheet)
    {
        foreach (var index in Indices)
        {
            if (sheet.ToSheetColumn(index) is var column and not -1)
            {
                return column;
            }
        }

        return -1;
    }

    public override int[]? FindColumns(SheetView sheet)
    {
        var columns = new int[Indices.Count];
        for (int i = 0; i < columns.Length; i++)
        {
            columns[i] = sheet.ToSheetColumn(Indices[i]);
            if (columns[i] == -1)
            {
                return null;
            }
        }

        return columns;
    }

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, rowIndex, member, Indices.Count == 1 ? $"column at index {Indices[0]}" : $"columns at indices {string.Join(", ", Indices)}");

    public override string Describe(SheetView sheet)
        => $"columns {string.Join(", ", Indices.Select(i => ExcelRange.GetColumnLetters(sheet.ColumnOffset + i)))}";
}

/// <summary>
/// The columns whose names match a predicate or a regular expression.
/// </summary>
internal sealed class MatchingColumns : Columns
{
    private readonly Func<string, bool> _matches;

    public MatchingColumns(Func<string, bool> predicate)
    {
        ThrowHelpers.ThrowIfNull(predicate, nameof(predicate));
        _matches = predicate;
    }

    public MatchingColumns(Regex regex)
    {
        ThrowHelpers.ThrowIfNull(regex, nameof(regex));
        _matches = regex.IsMatch;
    }

    public override int FindColumn(SheetView sheet)
    {
        var (start, end) = ColumnUtilities.GetColumnRange(sheet);
        for (var column = start; column < end; column++)
        {
            if (Matches(sheet, column))
            {
                return column;
            }
        }

        return -1;
    }

    public override int[]? FindColumns(SheetView sheet)
    {
        var columns = new List<int>();
        var (start, end) = ColumnUtilities.GetColumnRange(sheet);
        for (var column = start; column < end; column++)
        {
            if (Matches(sheet, column))
            {
                columns.Add(column);
            }
        }

        return columns.Count == 0 ? null : [.. columns];
    }

    private bool Matches(SheetView sheet, int column)
        => _matches((sheet.Heading ?? throw NoHeading(sheet)).GetColumnName(column));

    public override bool LooksUpNames => true;

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
    {
        if (sheet.Heading is not { } heading)
        {
            return Missing(sheet, rowIndex, member, null);
        }

        // The names of the matching columns, which a row may be too short to have.
        var names = new List<string>();
        var (start, end) = ColumnUtilities.GetColumnRange(sheet);
        for (var column = start; column < end; column++)
        {
            if (Matches(sheet, column))
            {
                names.Add(heading.GetColumnName(column)!);
            }
        }

        return names.Count == 0
            ? new($"Could not read value for member \"{member?.Name}\" (no columns matching)", sheet, rowIndex, -1)
            : Missing(sheet, rowIndex, member, DescribeNames(names));
    }
}

/// <summary>
/// One cell whose text is split into the values of a collection, such as "1,2,3": the cell is in
/// the first column <see cref="Cell"/> finds.
/// </summary>
internal sealed class SplitColumns(Columns cell, CellSplitter splitter) : Columns
{
    public Columns Cell { get; } = cell;

    public CellSplitter Splitter { get; } = splitter;

    public override int FindColumn(SheetView sheet) => Cell.FindColumn(sheet);

    public override bool LooksUpNames => Cell.LooksUpNames;

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Cell switch
        {
            ColumnNameColumns or FirstColumnNameColumns => Missing(sheet, rowIndex, member, $"column \"{Cell.ColumnName}\""),
            ColumnIndexColumns index => Missing(sheet, rowIndex, member, $"column at index {index.Index}"),
            _ => Missing(sheet, rowIndex, member, null),
        };

    public override string Describe(SheetView sheet)
        => Cell switch
        {
            ColumnNameColumns or FirstColumnNameColumns => $"one of columns \"{Cell.ColumnName}\"",
            ColumnIndexColumns index => $"columns {ExcelRange.GetColumnLetters(sheet.ColumnOffset + index.Index)}",
            _ => base.Describe(sheet),
        };
}

/// <summary>
/// The columns an <see cref="IColumnSelector"/> chooses. In a region, columns outside it are ignored.
/// </summary>
internal sealed class SelectorColumns(IColumnSelector selector) : Columns
{
    public IColumnSelector Selector { get; } = selector;

    public override int FindColumn(SheetView sheet)
    {
        var columns = Select(sheet);
        return columns.Length == 0 ? -1 : columns[0];
    }

    public override int[]? FindColumns(SheetView sheet)
    {
        var columns = Select(sheet);
        return columns.Length == 0 ? null : columns;
    }

    // The chosen columns that are in the sheet (or region).
    private int[] Select(SheetView sheet)
    {
        var selected = Selector.SelectColumns(new ColumnSelectorContext(sheet)) ?? throw new InvalidOperationException($"The column selector \"{Selector.GetType()}\" returned null.");
        var end = Math.Min(sheet.ColumnEnd, sheet.NumberOfColumns);
        var columns = new List<int>(selected.Count);
        for (int i = 0; i < selected.Count; i++)
        {
            if (selected[i] >= sheet.ColumnOffset && selected[i] < end)
            {
                columns.Add(selected[i]);
            }
        }

        return [.. columns];
    }

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, rowIndex, member, Selector.Describe(new ColumnSelectorContext(sheet)));

    public override string Describe(SheetView sheet) => Selector.Describe(new ColumnSelectorContext(sheet));
}

/// <summary>
/// Every column with a name in the heading: the default for a dictionary of a row's columns.
/// </summary>
internal sealed class AllColumns : Columns
{
    public static AllColumns Instance { get; } = new();

    private AllColumns()
    {
    }

    public override int FindColumn(SheetView sheet) => throw new NotSupportedException();

    public override int[] FindColumns(SheetView sheet)
    {
        var heading = sheet.Heading ?? throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");

        // Use the positions of the columns rather than looking them up by name, so that columns
        // with the same name are each read.
        var columns = new List<int>();
        var names = heading.ColumnNames;
        for (int i = 0; i < names.Count; i++)
        {
            if (!string.IsNullOrWhiteSpace(names[i]))
            {
                columns.Add(i + heading.ColumnOffset);
            }
        }

        return [.. columns];
    }
}

/// <summary>
/// The columns with a name that no other member of a class map reads: see
/// <see cref="ExcelExtraColumnsAttribute"/>. When rows are read, the class map is the one rows
/// are mapped with (see <see cref="SheetView.RootMap"/>), as a nested class map can be shared.
/// </summary>
internal sealed class ExtraColumns : Columns
{
    public ExtraColumns(ExcelClassMap classMap)
    {
        ThrowHelpers.ThrowIfNull(classMap, nameof(classMap));
        ClassMap = classMap;
    }

    public ExcelClassMap ClassMap { get; }

    public override int FindColumn(SheetView sheet) => throw new NotSupportedException();

    public override int[] FindColumns(SheetView sheet)
    {
        var heading = sheet.Heading ?? throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Extra columns can only be read from sheets with a heading.");

        var usedColumns = new HashSet<int>();
        AddUsedColumns(sheet.RootMap ?? ClassMap, sheet, usedColumns);

        var columns = new List<int>();
        var names = heading.ColumnNames;
        for (int i = 0; i < names.Count; i++)
        {
            var column = i + heading.ColumnOffset;
            if (!string.IsNullOrWhiteSpace(names[i]) && !usedColumns.Contains(column))
            {
                columns.Add(column);
            }
        }

        return [.. columns];
    }

    private static void AddUsedColumns(IMap map, SheetView sheet, HashSet<int> usedColumns)
        => MapTree.VisitCellMaps(map, null, (cellMap, _) =>
        {
            switch (cellMap)
            {
                case IValueMap valueMap:
                    AddUsedColumn(valueMap.Columns, sheet, usedColumns);
                    break;

                case ICellsMap { Columns: ExtraColumns }:
                    // Another extra columns member: it doesn't claim any columns.
                    break;

                case ICellsMap cellsMap:
                    try
                    {
                        if (cellsMap.Columns.FindColumns(sheet) is { } columns)
                        {
                            usedColumns.UnionWith(columns);
                        }
                    }
                    catch (ExcelMappingException)
                    {
                        // The member's columns can't be found, so it doesn't read any of the sheet's columns.
                    }

                    break;
            }
        });

    private static void AddUsedColumn(Columns columns, SheetView sheet, HashSet<int> usedColumns)
    {
        try
        {
            if (columns.FindColumn(sheet) is var column and not -1)
            {
                usedColumns.Add(column);
            }
        }
        catch (ExcelMappingException)
        {
            // The member's column can't be found (for example an optional column that is
            // missing), so it doesn't read any of the sheet's columns.
        }
    }
}

/// <summary>
/// The regular expression given to <see cref="ExcelColumnMatchingAttribute"/> or
/// <see cref="ExcelColumnsMatchingAttribute"/>, whose <c>Type</c> is this type.
/// </summary>
internal sealed class RegexColumnMatcher(Regex regex)
{
    public Regex Regex { get; } = regex;
}
