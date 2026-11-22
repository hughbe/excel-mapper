using System.Globalization;
using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// A sheet reads the importer's settings when it starts reading, so changing them part way
/// through doesn't change how the rest of the sheet is read.
/// </summary>
public class MapConfigurationChangesTests
{
    [Fact]
    public void ReadRows_CultureChangedWhileReading_KeepsTheSheetsCulture()
    {
        using var importer = Helpers.GetCsvImporter("Amount\n\"1,5\"\n\"2,5\"\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("de-DE");
        var sheet = importer.ReadSheet();
        var amounts = new List<double>();
        foreach (var row in sheet.ReadRows<Row>())
        {
            amounts.Add(row.Amount);

            // In the invariant culture, "2,5" would be 25.
            importer.Configuration.Culture = CultureInfo.InvariantCulture;
        }

        Assert.Equal([1.5, 2.5], amounts);
    }

    [Fact]
    public void ReadRows_CultureChangedBetweenSheets_NextSheetUsesTheNewCulture()
    {
        using var importer = Helpers.GetCsvImporter("Amount\n\"1,5\"\n");
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("de-DE");
        Assert.Equal(1.5, importer.ReadSheet().ReadRows<Row>().Single().Amount);

        importer.Configuration.Culture = CultureInfo.InvariantCulture;
        var region = importer.ReadSheet(0).Region();
        Assert.Equal(15, region.ReadRows<Row>().Single().Amount);
    }

    public class Row
    {
        public double Amount { get; set; }
    }
}
