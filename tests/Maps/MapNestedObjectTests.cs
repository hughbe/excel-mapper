using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class MapNestedObjectTests
{
    [Fact]
    public void ReadRow_AutoMappedObject_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("NestedObjects.xlsx");

        ExcelSheet sheet = importer.ReadSheet();
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

        ExcelSheet sheet = importer.ReadSheet();
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

        ExcelSheet sheet = importer.ReadSheet();
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

        ExcelSheet sheet = importer.ReadSheet();
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

    [Fact]
    public void ReadRow_MapNestedList_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("NestedList.xlsx");
        importer.Configuration.RegisterClassMap<NestedListParentClassMap>();

        ExcelSheet sheet = importer.ReadSheet();
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

    private class NestedListParentClassMap : ExcelClassMap<NestedListParentClass>
    {
        public NestedListParentClassMap()
        {
            Map(v => v.Name);
            Map(v => v.Address);

            var member = typeof(NestedListParentClass).GetProperty(nameof(NestedListParentClass.BusinessHours))!;
            Properties.Add(new ExcelPropertyMap<List<BusinessHours>>(member, new BusinessHoursMap()));
        }
    }

    private class BusinessHoursMap : IMap
    {
        private int _previousRowIndex = -1;
        private int _currentIndex = 0;

        public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
        {
            // Note: only works for 2 days (Monday, Tuesday) as written, but easy to extend.
            var result = new List<BusinessHours>();
            for (int i = 0; i < 2; i++)
            {
                if (_previousRowIndex != rowIndex)
                {
                    _previousRowIndex = rowIndex;
                    _currentIndex = 0;
                }

                string prefix;
                switch (_currentIndex)
                {
                    case 0:
                        prefix = "Monday";
                        break;
                    case 1:
                        prefix = "Tuesday";
                        break;
                    default:
                        throw new NotImplementedException();
                }

                // Onto the next column.
                _currentIndex++;

                // Format: "<DayOfWeek>DayLabel"
                var labelReaderFactory = new ColumnNameReaderFactory(prefix + "Label");
                var startTimeReaderFactory = new ColumnNameReaderFactory(prefix + "Open");
                var endTimeReaderFactory = new ColumnNameReaderFactory(prefix + "Close");

                var labelReader = labelReaderFactory.GetCellReader(sheet)!;
                var startTimeReader = startTimeReaderFactory.GetCellReader(sheet)!;
                var endTimeReader = endTimeReaderFactory.GetCellReader(sheet)!;

                if (!labelReader.TryGetValue(reader, false, out ReadCellResult labelResult) ||
                    !startTimeReader.TryGetValue(reader, false,  out ReadCellResult startTimeResult) ||
                    !endTimeReader.TryGetValue(reader, false, out ReadCellResult endTimeResult))
                {
                    throw new InvalidOperationException("No such column");
                }

                result.Add(new BusinessHours
                {
                    DayLabel = labelResult.StringValue,
                    StartTime = startTimeResult.StringValue,
                    EndTime = endTimeResult.StringValue
                });            
            }

            value = result;
            return true;
        }
    }
}
