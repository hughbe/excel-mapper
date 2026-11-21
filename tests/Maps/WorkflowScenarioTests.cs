using System.ComponentModel.DataAnnotations;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelMapper.Readers;

namespace ExcelMapper.Tests;

/// <summary>
/// Workflow and configuration scenarios: several sheets, several importers, re-registering maps,
/// configuration options and CSV quirks.
/// </summary>
public class WorkflowScenarioTests
{
    private const string PeopleCsv = "Name,Age\nAlice,30\nBob,25\nCarol,41\n";

    public record Person(string Name, int Age);

    public class PersonClass
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // ---------- Sheets and importers ----------

    [Fact]
    public void ReadSheets_ReadRowsOfEachSheet_EnumeratesAllSheets()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var names = new List<string>();
        foreach (var sheet in importer.ReadSheets())
        {
            names.Add(sheet.Name);
            if (sheet.Name == "Primitives")
            {
                Assert.NotEmpty(sheet.ReadRows<string[]>().ToArray());
            }
        }

        Assert.Equal(["Primitives", "Empty", "Third Sheet"], names);
    }

    [Fact]
    public void ReadRow_InterleavedSheets_EachSheetKeepsItsPosition()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var first = importer.ReadSheet("Primitives");
        var third = importer.ReadSheet("Third Sheet");

        var firstRow1 = first.ReadRow<string[]>();
        third.ReadRow<string[]>();
        var firstRow2 = first.ReadRow<string[]>();

        Assert.NotEqual(firstRow1, firstRow2);
        Assert.Equal(2, first.CurrentRowIndex);
        Assert.Equal(1, third.CurrentRowIndex);
    }

    [Fact]
    public void Dispose_Twice_DoesNotThrow()
    {
        var importer = Helpers.GetCsvImporter(PeopleCsv);
        importer.Dispose();
        importer.Dispose();
    }

    // ---------- Sharing and registering maps ----------

    [Fact]
    public void RegisterClassMap_AfterAutoMapping_ReplacesAutoMap()
    {
        using var importer = Helpers.GetCsvImporter("Full Name,Age\nAlice,30\nBob,25\n");
        var sheet = importer.ReadSheet();

        // Auto-mapping fails as there is no "Name" column, but registers the auto-map.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<PersonClass>());

        importer.Configuration.RegisterClassMap<PersonClass>(c =>
        {
            c.Map(p => p.Name).WithColumnName("Full Name");
            c.Map(p => p.Age);
        });

        var row = sheet.ReadRow<PersonClass>();
        Assert.Equal("Bob", row.Name);
        Assert.Equal(25, row.Age);

        // A map registered by the user can't be replaced.
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<PersonClass>(c => c.Map(p => p.Name)));
    }

    [Fact]
    public void RegisterClassMap_SameMapInTwoImportersWithDifferentColumnOrders_Success()
    {
        var map = new ExcelClassMap<PersonClass>();
        map.Map(p => p.Name);
        map.Map(p => p.Age);

        using var first = Helpers.GetCsvImporter(PeopleCsv);
        first.Configuration.RegisterClassMap(map);
        using var second = Helpers.GetCsvImporter("Age,Name\n1,Zed\n");
        second.Configuration.RegisterClassMap(map);

        Assert.Equal("Alice", first.ReadSheet().ReadRows<PersonClass>().First().Name);
        Assert.Equal("Zed", second.ReadSheet().ReadRows<PersonClass>().Single().Name);
    }

    public class CollectionsClass
    {
        public List<int> Scores { get; set; } = default!;
        public Dictionary<string, string> Extra { get; set; } = default!;
        public int[] Indexed { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_SharedMapInParallelImporters_Success()
    {
        var map = new ExcelClassMap<CollectionsClass>();
        map.Map(p => p.Scores);
        map.Map(p => p.Extra).WithColumnNames("A", "B");
        map.Map(p => p.Indexed[0]).WithColumnName("A");
        map.Map(p => p.Indexed[1]).WithColumnName("B");

        var csv = "Scores,A,B\n" + string.Join("\n", Enumerable.Repeat("\"1,2,3\",4,5", 200)) + "\n";
        var failures = 0;
        Parallel.For(0, 32, _ =>
        {
            try
            {
                using var importer = Helpers.GetCsvImporter(csv);
                importer.Configuration.RegisterClassMap(map);
                foreach (var row in importer.ReadSheet().ReadRows<CollectionsClass>())
                {
                    if (!row.Scores.SequenceEqual([1, 2, 3]) || row.Extra.Count != 2 || !row.Indexed.SequenceEqual([4, 5]))
                    {
                        Interlocked.Increment(ref failures);
                    }
                }
            }
            catch (Exception)
            {
                Interlocked.Increment(ref failures);
            }
        });

        Assert.Equal(0, failures);
    }

    [Fact]
    public void ReadRows_SharedMap_DoesNotKeepSheetsAlive()
    {
        var map = new ExcelClassMap<CollectionsClass>();
        map.Map(p => p.Scores);

        var sheetReference = ReadWithMap(map);
        for (int i = 0; i < 5 && sheetReference.IsAlive; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        Assert.False(sheetReference.IsAlive);
        GC.KeepAlive(map);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static WeakReference ReadWithMap(ExcelClassMap<CollectionsClass> map)
    {
        using var importer = Helpers.GetCsvImporter("Scores\n\"1,2\"\n");
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();
        Assert.Equal([1, 2], sheet.ReadRows<CollectionsClass>().Single().Scores);
        return new WeakReference(sheet);
    }

    public class NameClass
    {
        public string Name { get; set; } = "";
    }

    [Fact]
    public void ReaderFactory_ChangedAfterReading_UsesNewColumn()
    {
        using var importer = Helpers.GetCsvImporter("A,B\na1,b1\na2,b2\n");
        var map = new ExcelClassMap<NameClass>();
        var nameMap = map.Map(p => p.Name).WithColumnName("A");
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();

        Assert.Equal("a1", sheet.ReadRow<NameClass>().Name);

        nameMap.WithColumnName("B");
        Assert.Equal("b2", sheet.ReadRow<NameClass>().Name);
    }

    // ---------- Configuration ----------

    [Fact]
    public void ReadRows_RangeOfRecords_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var sheet = importer.ReadSheet();
        Assert.Equal([new Person("Bob", 25), new Person("Carol", 41)], sheet.ReadRows<Person>(2, 2));

        // Earlier rows can be read again.
        Assert.Equal([new Person("Alice", 30)], sheet.ReadRows<Person>(1, 1));
    }

    [Fact]
    public void ReadRows_SkipBlankLinesRecords_Success()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,30\n,\nBob,25\n");
        importer.Configuration.SkipBlankLines = true;
        Assert.Equal([new Person("Alice", 30), new Person("Bob", 25)], importer.ReadSheet().ReadRows<Person>());
    }

    public record ValidatedRecord([property: Range(0, 120)] int Age, [property: Required] string Name);

    [Fact]
    public void ReadRows_ValidateDataAnnotationsOnRecord_ThrowsValidationException()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,30\nBob,500\n");
        importer.Configuration.ValidateDataAnnotations = true;
        var sheet = importer.ReadSheet();

        Assert.Equal(new ValidatedRecord(30, "Alice"), sheet.ReadRow<ValidatedRecord>());
        Assert.Throws<ValidationException>(() => sheet.ReadRow<ValidatedRecord>());
    }

    [Fact]
    public void ReadRows_HeadingIndexRecords_Success()
    {
        using var importer = Helpers.GetCsvImporter("Report\nName,Age\nAlice,30\n");
        var sheet = importer.ReadSheet();
        sheet.HeadingIndex = 1;
        Assert.Equal(new Person("Alice", 30), sheet.ReadRows<Person>().Single());
    }

    // ---------- CSV and heading quirks ----------

    [Theory]
    [InlineData("Name;Age\nAlice;30\n")]
    [InlineData("Name\tAge\nAlice\t30\n")]
    [InlineData("Name|Age\nAlice|30\n")]
    public void ReadRows_CsvSeparators_Success(string csv)
    {
        using var importer = Helpers.GetCsvImporter(csv);
        Assert.Equal(new Person("Alice", 30), importer.ReadSheet().ReadRows<Person>().Single());
    }

    [Fact]
    public void ReadRows_WhitespaceAroundHeadingsAndNumbers_Success()
    {
        using var importer = Helpers.GetCsvImporter(" Name , Age \nAlice, 30 \n");
        Assert.Equal(new Person("Alice", 30), importer.ReadSheet().ReadRows<Person>().Single());
    }

    [Fact]
    public void ReadRows_DuplicateColumns_UsesFirst()
    {
        using var importer = Helpers.GetCsvImporter("Name,Name,Age\nFirst,Second,1\n");
        Assert.Equal("First", importer.ReadSheet().ReadRows<Person>().Single().Name);
    }

    [Fact]
    public void ReadRows_ExpandoObject_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        dynamic row = importer.ReadSheet().ReadRows<ExpandoObject>().First();
        Assert.Equal("Alice", (string)row.Name);
        Assert.Equal("30", (string)row.Age);
    }

    [Fact]
    public void ReadRows_HeaderOnlyCsv_ReturnsEmpty()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\n");
        Assert.Empty(importer.ReadSheet().ReadRows<Person>());
    }

    public record CommonTypes(Guid Id, Uri Site, Version Version, TimeSpan Duration, decimal Money, bool Flag);

    [Fact]
    public void ReadRows_CommonTypesInRecord_Success()
    {
        using var importer = Helpers.GetCsvImporter($"Id,Site,Version,Duration,Money,Flag\n{Guid.Empty},https://example.org,1.2.3,01:02:03,1.5,true\n");
        var row = importer.ReadSheet().ReadRows<CommonTypes>().Single();
        Assert.Equal(new CommonTypes(Guid.Empty, new Uri("https://example.org"), new Version(1, 2, 3), new TimeSpan(1, 2, 3), 1.5m, true), row);
    }

    // ---------- Should fail clearly ----------

    [Fact]
    public void ReadRows_HeadingIndexPastEnd_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var sheet = importer.ReadSheet();
        sheet.HeadingIndex = 50;
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Person>().ToArray());
    }

    [Fact]
    public void ReadSheet_NoMoreSheets_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet());
    }

    [Fact]
    public void ReadRows_RangePastEnd_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<Person>(2, 10).ToArray());
    }

    [Fact]
    public void HasHeading_SetAfterReading_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var sheet = importer.ReadSheet();
        sheet.ReadRows<Person>().First();
        Assert.Throws<InvalidOperationException>(() => sheet.HasHeading = false);
        Assert.Throws<InvalidOperationException>(() => sheet.DataRange = new ExcelRange("A:B"));
    }

    [Fact]
    public void Ctor_NotAnExcelFile_Throws()
    {
        Assert.ThrowsAny<Exception>(() => new ExcelImporter(new MemoryStream([1, 2, 3, 4, 5, 6, 7, 8])));
    }
}
