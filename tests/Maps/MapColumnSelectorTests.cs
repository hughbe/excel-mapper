using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// Choosing columns with an <see cref="IColumnSelector"/>, set with WithColumns or the column
/// matching attributes.
/// </summary>
public class MapColumnSelectorTests
{
    [Fact]
    public void ReadRows_SingleValue_ReadsFirstColumnChosen()
    {
        using var importer = Helpers.GetCsvImporter("Name,Amount (GBP),Amount (USD)\na,1,2\nb,3,4\n");
        importer.Configuration.RegisterClassMap<AmountRow>(c => c.Map(r => r.Amount).WithColumns(new PrefixSelector("Amount")));
        Assert.Equal([1, 3], importer.ReadSheet().ReadRows<AmountRow>().Select(r => r.Amount));
    }

    [Fact]
    public void ReadRows_Collection_ReadsColumnsInOrderChosen()
    {
        using var importer = Helpers.GetCsvImporter("Q1,Name,Q2,Q3\n1,a,2,3\n");
        importer.Configuration.RegisterClassMap<QuartersRow>(c => c.Map(r => r.Quarters).WithColumns(new ReversePrefixSelector("Q")));
        Assert.Equal([3, 2, 1], importer.ReadSheet().ReadRows<QuartersRow>().Single().Quarters);
    }

    [Fact]
    public void ReadRows_Dictionary_KeysAreColumnNames()
    {
        using var importer = Helpers.GetCsvImporter("Q1,Name,Q2\n1,a,2\n");
        importer.Configuration.RegisterClassMap<DictionaryRow>(c => c.Map(r => r.Quarters).WithColumns(new PrefixSelector("Q")));
        Assert.Equal(new Dictionary<string, int> { ["Q1"] = 1, ["Q2"] = 2 }, importer.ReadSheet().ReadRows<DictionaryRow>().Single().Quarters);
    }

    [Fact]
    public void ReadRows_Attributes_UseSelector()
    {
        using var importer = Helpers.GetCsvImporter("Name,Amount (GBP),Q1,Q2\na,1,2,3\n");
        var row = importer.ReadSheet().ReadRows<AttributeRow>().Single();
        Assert.Equal(1, row.Amount);
        Assert.Equal([2, 3], row.Quarters);
    }

    [Fact]
    public void ReadRows_NoColumnChosen_ThrowsWithDescription()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        importer.Configuration.RegisterClassMap<AmountRow>(c => c.Map(r => r.Amount).WithColumns(new PrefixSelector("Amount")));
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<AmountRow>().ToArray());
        Assert.Contains("columns starting with \"Amount\"", exception.Message);
    }

    [Fact]
    public void ReadRows_NoColumnChosenOptional_KeepsDefault()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        importer.Configuration.RegisterClassMap<AmountRow>(c => c.Map(r => r.Amount).WithColumns(new PrefixSelector("Amount")).MakeOptional());
        Assert.Equal(0, importer.ReadSheet().ReadRows<AmountRow>().Single().Amount);
    }

    [Fact]
    public void TryValidateHeading_NoColumnChosen_ReportsDescription()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        importer.Configuration.RegisterClassMap<AmountRow>(c => c.Map(r => r.Amount).WithColumns(new PrefixSelector("Amount")));
        Assert.False(importer.ReadSheet().TryValidateHeading<AmountRow>(out var problems));
        Assert.Contains("columns starting with \"Amount\"", Assert.Single(problems));
    }

    [Fact]
    public void ReadRows_ManyRows_SelectsOncePerSheet()
    {
        var selector = new CountingSelector();
        using var importer = Helpers.GetCsvImporter("Amount\n1\n2\n3\n");
        importer.Configuration.RegisterClassMap<AmountRow>(c => c.Map(r => r.Amount).WithColumns(selector));
        Assert.Equal([1, 2, 3], importer.ReadSheet().ReadRows<AmountRow>().Select(r => r.Amount));
        Assert.Equal(1, selector.Calls);
    }

    [Fact]
    public void ReadRows_SheetsWithDifferentLayouts_SelectsForEachSheet()
    {
        var map = new ExcelClassMap<AmountRow>();
        map.Map(r => r.Amount).WithColumns(new PrefixSelector("Amount"));
        using var first = Helpers.GetCsvImporter("Amount,Name\n1,a\n");
        using var second = Helpers.GetCsvImporter("Name,Amount\na,2\n");
        first.Configuration.RegisterClassMap(map);
        second.Configuration.RegisterClassMap(map);
        Assert.Equal(1, first.ReadSheet().ReadRows<AmountRow>().Single().Amount);
        Assert.Equal(2, second.ReadSheet().ReadRows<AmountRow>().Single().Amount);
    }

    [Fact]
    public void ReadRows_Region_IgnoresColumnsOutsideRegion()
    {
        using var importer = Helpers.GetCsvImporter("Amount,Name,Amount\n1,a,2\n");
        importer.Configuration.RegisterClassMap<AmountRow>(c => c.Map(r => r.Amount).WithColumns(new PositionsSelector(0, 2)));
        var sheet = importer.ReadSheet();
        Assert.Equal(1, sheet.ReadRows<AmountRow>().Single().Amount);
        Assert.Equal(2, sheet.Region("B1:C2").ReadRows<AmountRow>().Single().Amount);
    }

    [Fact]
    public void ReadRows_WithoutHeading_SelectsByPosition()
    {
        using var importer = Helpers.GetCsvImporter("1,2\n3,4\n");
        importer.Configuration.RegisterClassMap<AmountRow>(c => c.Map(r => r.Amount).WithColumns(new PositionsSelector(1)));
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        Assert.Equal([2, 4], sheet.ReadRows<AmountRow>().Select(r => r.Amount));
    }

    [Fact]
    public void ReadRows_ExtraColumns_ExcludeSelectedColumns()
    {
        using var importer = Helpers.GetCsvImporter("Name,Amount (GBP),Notes\na,1,n\n");
        importer.Configuration.RegisterClassMap<ExtraRow>(c =>
        {
            c.Map(r => r.Name);
            c.Map(r => r.Amount).WithColumns(new PrefixSelector("Amount"));
            c.MapExtraColumns(r => r.Other);
        });
        Assert.Equal(new Dictionary<string, string> { ["Notes"] = "n" }, importer.ReadSheet().ReadRows<ExtraRow>().Single().Other);
    }

    [Fact]
    public void ReadRows_SelectorReturnsNull_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetCsvImporter("Amount\n1\n");
        importer.Configuration.RegisterClassMap<AmountRow>(c => c.Map(r => r.Amount).WithColumns(new NullSelector()));
        var sheet = importer.ReadSheet();
        Assert.Throws<InvalidOperationException>(() => sheet.ReadRows<AmountRow>().ToArray());
    }

    public class AmountRow
    {
        public int Amount { get; set; }
    }

    public class QuartersRow
    {
        public int[] Quarters { get; set; } = [];
    }

    public class DictionaryRow
    {
        public Dictionary<string, int> Quarters { get; set; } = [];
    }

    public class AttributeRow
    {
        [ExcelColumnMatching(typeof(PrefixSelector), ConstructorArguments = ["Amount"])]
        public int Amount { get; set; }

        [ExcelColumnsMatching(typeof(PrefixSelector), ConstructorArguments = ["Q"])]
        public int[] Quarters { get; set; } = [];
    }

    public class ExtraRow
    {
        public string? Name { get; set; }
        public int Amount { get; set; }
        public Dictionary<string, string> Other { get; set; } = [];
    }

    public class PrefixSelector(string prefix) : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context)
            => context.Heading!.ColumnNames.Where(n => n.StartsWith(prefix, StringComparison.Ordinal)).Select(context.Heading.GetColumnIndex).ToArray();

        public string Describe(ColumnSelectorContext context) => $"columns starting with \"{prefix}\"";
    }

    public class ReversePrefixSelector(string prefix) : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context) => new PrefixSelector(prefix).SelectColumns(context).Reverse().ToArray();

        public string Describe(ColumnSelectorContext context) => $"columns starting with \"{prefix}\", last first";
    }

    public class PositionsSelector(params int[] positions) : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context) => positions;

        public string Describe(ColumnSelectorContext context) => $"columns {string.Join(", ", positions)}";
    }

    private sealed class CountingSelector : IColumnSelector
    {
        public int Calls { get; private set; }

        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context)
        {
            Calls++;
            return [0];
        }

        public string Describe(ColumnSelectorContext context) => "the first column";
    }

    private sealed class NullSelector : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ColumnSelectorContext context) => null!;

        public string Describe(ColumnSelectorContext context) => "nothing";
    }
}
