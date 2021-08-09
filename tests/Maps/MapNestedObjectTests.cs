using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq.Expressions;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapNestedObjectTests
    {
        [Fact]
        public void ReadRow_AutoMappedObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("NestedObjects.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            NestedObjectValue row1 = sheet.ReadRow<NestedObjectValue>();
            Assert.Equal("a", row1.SubValue1.StringValue);
            Assert.Equal(new string[] { "a", "b" }, row1.SubValue1.SplitStringValue);
            Assert.Equal(1, row1.SubValue2.IntValue);
            Assert.Equal(10, row1.SubValue2.SubValue.SubInt);
            Assert.Equal("c", row1.SubValue2.SubValue.SubString);
        }

        private class NestedObjectValue
        {
            public SubValue1 SubValue1 { get; set; }
            public SubValue2 SubValue2 { get; set; }
        }

        private class SubValue1
        {
            public string StringValue { get; set; }
            public string[] SplitStringValue { get; set; }
        }

        private class SubValue2
        {
            public int IntValue { get; set; }
            public SubValue3 SubValue { get; set; }
        }

        private class SubValue3
        {
            public string SubString { get; set; }
            public int SubInt { get; set; }
        }

        [Fact]
        public void ReadRow_CustomMappedObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("NestedObjects.xlsx");
            importer.Configuration.RegisterClassMap<ObjectValueCustomClassMapMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            NestedObjectValue row1 = sheet.ReadRow<NestedObjectValue>();
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

            NestedObjectValue row1 = sheet.ReadRow<NestedObjectValue>();
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
        public void ReadRow_MapNestedList_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("NestedList.xlsx");
            importer.Configuration.RegisterClassMap<NestedListParentClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            NestedListParentClass row1 = sheet.ReadRow<NestedListParentClass>();
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

        public class NestedListParentClass
        {
            public string Name { get; set; }
            public string Address { get; set; }
            public List<BusinessHours> BusinessHours { get; set; }
        }

        public class BusinessHours
        {
            public string DayLabel { get; set; }
            public string StartTime { get; set; }
            public string EndTime { get; set; }
        }

        public class NestedListParentClassMap : ExcelClassMap<NestedListParentClass>
        {
            public NestedListParentClassMap()
            {
                Map(v => v.Name);
                Map(v => v.Address);
                Expression<Func<NestedListParentClass, List<BusinessHours>>> expression = v => v.BusinessHours;
                var businessMap = new ExcelPropertyMap(GetMemberExpression(expression).Member, new BusinessHoursMap());
                AddMap(businessMap, v => v.BusinessHours);
            }
        }

        public class BusinessHoursMap : IMap
        {
            private int _previousRowIndex = -1;
            private int _currentIndex = 0;

            public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo member, out object value)
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
                    var labelReader = new ColumnNameValueReader(prefix + "Label");
                    var startTimeReader = new ColumnNameValueReader(prefix + "Open");
                    var endTimeReader = new ColumnNameValueReader(prefix + "Close");
                    if (!labelReader.TryGetValue(sheet, rowIndex, reader, out ReadCellValueResult labelResult) ||
                        !startTimeReader.TryGetValue(sheet, rowIndex, reader, out ReadCellValueResult startTimeResult) ||
                        !endTimeReader.TryGetValue(sheet, rowIndex, reader, out ReadCellValueResult endTimeResult))
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
}
