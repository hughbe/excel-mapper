using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Tests;

public class MapNestedObjectTests
{
    [Fact]
    public void ReadRow_AutoMappedObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("NestedObjects.xlsx");

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NestedObjectValue>();
        Assert.Equal("a", row1.SubValue1.StringValue);
        Assert.Equal(new string[] { "a", "b" }, row1.SubValue1.SplitStringValue);
        Assert.Equal(1, row1.SubValue2.IntValue);
        Assert.Equal(10, row1.SubValue2.SubValue.SubInt);
        Assert.Equal("c", row1.SubValue2.SubValue.SubString);
    }

    private class NestedObjectValue
    {
        public SubValue1 SubValue1 { get; set; } = default!;
        public SubValue2 SubValue2 { get; set; } = default!;
    }

    private class SubValue1
    {
        public string StringValue { get; set; } = default!;
        public string[] SplitStringValue { get; set; } = default!;
    }

    private class SubValue2
    {
        public int IntValue { get; set; }
        public SubValue3 SubValue { get; set; } = default!;
    }

    private class SubValue3
    {
        public string SubString { get; set; } = default!;
        public int SubInt { get; set; }
    }

    [Fact]
    public void ReadRow_CustomMappedObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("NestedObjects.xlsx");
        importer.Configuration.RegisterClassMap<ObjectValueCustomClassMapMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NestedObjectValue>();
        Assert.Equal("a", row1.SubValue1.StringValue);
        Assert.Equal(new string[] { "a", "b" }, row1.SubValue1.SplitStringValue);
        Assert.Equal(1, row1.SubValue2.IntValue);
        Assert.Equal(10, row1.SubValue2.SubValue.SubInt);
        Assert.Equal("c", row1.SubValue2.SubValue.SubString);
    }

    private class ObjectValueCustomClassMapMap : ExcelClassMap<NestedObjectValue>
    {
        public ObjectValueCustomClassMapMap()
        {
            MapObject(p => p.SubValue1).WithClassMap(m =>
            {
                m.Map(s => s.StringValue);
                m.Map(s => s.SplitStringValue);
            });

            MapObject(p => p.SubValue2).WithClassMap(new SubValueMap());
        }
    }

    private class SubValueMap : ExcelClassMap<SubValue2>
    {
        public SubValueMap()
        {
            Map(s => s.IntValue);

            MapObject(s => s.SubValue);
        }
    }

    [Fact]
    public void ReadRow_CustomInnerObjectMap_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("NestedObjects.xlsx");
        importer.Configuration.RegisterClassMap<ObjectValueInnerMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NestedObjectValue>();
        Assert.Equal("a", row1.SubValue1.StringValue);
        Assert.Equal(new string[] { "a", "b" }, row1.SubValue1.SplitStringValue);
        Assert.Equal(1, row1.SubValue2.IntValue);
        Assert.Equal(10, row1.SubValue2.SubValue.SubInt);
        Assert.Equal("c", row1.SubValue2.SubValue.SubString);
    }

    private class ObjectValueInnerMap : ExcelClassMap<NestedObjectValue>
    {
        public ObjectValueInnerMap()
        {
            Map(p => p.SubValue1.StringValue);
            Map(p => p.SubValue1.SplitStringValue);
            Map(p => p.SubValue2.IntValue);
            Map(p => p.SubValue2.SubValue.SubInt);
            Map(p => p.SubValue2.SubValue.SubString);
        }
    }

    [Fact]
    public void ReadRow_MapNestedListIndexer_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("NestedList.xlsx");
        importer.Configuration.RegisterClassMap<NestedListParentIndexerClassMap>();

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<NestedListParentClass>();
        Assert.Equal("TheName", row1.Name);
        Assert.Equal("TheAddress", row1.Address);
        Assert.Equal(2, row1.BusinessHours.Count);
        Assert.Equal("TheMondayLabel", row1.BusinessHours[0].DayLabel);
        Assert.Equal("TheMondayOpen", row1.BusinessHours[0].StartTime);
        Assert.Equal("TheMondayClose", row1.BusinessHours[0].EndTime);
        Assert.Equal("TheTuesdayLabel", row1.BusinessHours[1].DayLabel);
        Assert.Equal("TheTuesdayOpen", row1.BusinessHours[1].StartTime);
        Assert.Equal("TheTuesdayClose", row1.BusinessHours[1].EndTime);
    }

    private class NestedListParentIndexerClassMap : ExcelClassMap<NestedListParentClass>
    {
        public NestedListParentIndexerClassMap()
        {
            Map(v => v.Name);
            Map(v => v.Address);

            Map(v => v.BusinessHours[0].DayLabel)
                .WithColumnName("MondayLabel");
            Map(v => v.BusinessHours[0].StartTime)
                .WithColumnName("MondayOpen");
            Map(v => v.BusinessHours[0].EndTime)
                .WithColumnName("MondayClose");
            Map(v => v.BusinessHours[1].DayLabel)
                .WithColumnName("TuesdayLabel");
            Map(v => v.BusinessHours[1].StartTime)
                .WithColumnName("TuesdayOpen");
            Map(v => v.BusinessHours[1].EndTime)
                .WithColumnName("TuesdayClose");
            // Note: only works for 2 days (Monday, Tuesday) as written, but easy to extend.
        }
    }

    private class NestedListParentClass
    {
        public string Name { get; set; } = default!;
        public string Address { get; set; } = default!;
        public List<BusinessHours> BusinessHours { get; set; } = default!;
    }

    private class BusinessHours
    {
        public string? DayLabel { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
    }

}
