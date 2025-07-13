using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualBasic;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapDateTimeTests
    {
        [Fact]
        public void ReadRow_AutoMappedDateTime_Success()
        {
            using var importer = Helpers.GetImporter("DateTimes.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            DateTimeValue row1 = sheet.ReadRow<DateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());
        }

        [Fact]
        public void ReadRow_AutoMappedNullableDateTime_Success()
        {
            using var importer = Helpers.GetImporter("DateTimes.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableDateTimeValue row1 = sheet.ReadRow<NullableDateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

            // Empty cell value.
            NullableDateTimeValue row5 = sheet.ReadRow<NullableDateTimeValue>();
            Assert.Null(row5.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDateTimeValue>());
        }

        [Fact]
        public void ReadRow_DefaultMappedDateTime_Success()
        {
            using var importer = Helpers.GetImporter("DateTimes.xlsx");
            importer.Configuration.RegisterClassMap<DefaultDateTimeClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            DateTimeValue row1 = sheet.ReadRow<DateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());
        }

        [Fact]
        public void ReadRow_DefaultMappedNullableDateTime_Success()
        {
            using var importer = Helpers.GetImporter("DateTimes.xlsx");
            importer.Configuration.RegisterClassMap<DefaultNullableDateTimeClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableDateTimeValue row1 = sheet.ReadRow<NullableDateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

            // Empty cell value.
            NullableDateTimeValue row5 = sheet.ReadRow<NullableDateTimeValue>();
            Assert.Null(row5.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDateTimeValue>());
        }

        [Fact]
        public void ReadRow_CustomMappedDateTime_Success()
        {
            using var importer = Helpers.GetImporter("DateTimes.xlsx");
            importer.Configuration.RegisterClassMap<CustomDateTimeClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            DateTimeValue row1 = sheet.ReadRow<DateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

            // Empty cell value.
            DateTimeValue row5 = sheet.ReadRow<DateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 20), row5.Value);

            // Invalid cell value.
            DateTimeValue row6 = sheet.ReadRow<DateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 21), row6.Value);
        }

        [Fact]
        public void ReadRow_CustomFormatsArrayDateTime_Success()
        {
            using var importer = Helpers.GetImporter("DateTimes.xlsx");
            importer.Configuration.RegisterClassMap<DateTimeFormatsArrayMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            CustomDateTimeValue row1 = sheet.ReadRow<CustomDateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 19), row1.CustomValue);

            CustomDateTimeValue row2 = sheet.ReadRow<CustomDateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 18), row2.CustomValue);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());
        }

        [Fact]
        public void ReadRow_CustomEnumerableFormatsDateTime_Success()
        {
            using var importer = Helpers.GetImporter("DateTimes.xlsx");
            importer.Configuration.RegisterClassMap<DateTimeEnumerableFormatsMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            CustomDateTimeValue row1 = sheet.ReadRow<CustomDateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 19), row1.CustomValue);

            CustomDateTimeValue row2 = sheet.ReadRow<CustomDateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 18), row2.CustomValue);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DateTimeValue>());
        }

        [Fact]
        public void ReadRow_CustomMappedNullableDateTime_Success()
        {
            using var importer = Helpers.GetImporter("DateTimes.xlsx");
            importer.Configuration.RegisterClassMap<CustomNullableDateTimeClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableDateTimeValue row1 = sheet.ReadRow<NullableDateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 19), row1.Value);

            // Empty cell value.
            NullableDateTimeValue row5 = sheet.ReadRow<NullableDateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 20), row5.Value);

            // Invalid cell value.
            NullableDateTimeValue row6 = sheet.ReadRow<NullableDateTimeValue>();
            Assert.Equal(new DateTime(2017, 07, 21), row6.Value);
        }

        [Fact]
        public void ReadRow_InvalidFormat_ThrowsErrorWithRightMessage()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using var importer = Helpers.GetImporter("DateTimes_Errors.xlsx");
            importer.Configuration.RegisterClassMap<DateTimeValueClassMap>();

            var sheet = importer.ReadSheet("Sheet1");
            sheet.HeadingIndex = 0;

            // Valid cell value.
            var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<DateTimeValueInt>(0, 2).ToArray());
            Assert.IsType<FormatException>(ex.InnerException);
            Assert.Equal("Invalid assigning \"1/1/2025 1:01:01 AM\" to member \"Value\" of type \"System.Int32\" in column \"Date\" on row 1 in sheet \"Sheet1\".", ex.Message);
        }

        private class DateTimeValueInt
        {
            public int Value { get; set; }
        }

        private class DateTimeValueClassMap : ExcelClassMap<DateTimeValueInt>
        {
            public DateTimeValueClassMap()
            {
                Map(m => m.Value).WithColumnIndex(0);
            }
        }

        private class DateTimeValue
        {
            public DateTime Value { get; set; }
        }

        private class CustomDateTimeValue
        {
            public DateTime CustomValue { get; set; }
        }

        private class NullableDateTimeValue
        {
            public DateTime? Value { get; set; }
        }

        private class DefaultDateTimeClassMap : ExcelClassMap<DateTimeValue>
        {
            public DefaultDateTimeClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomDateTimeClassMap : ExcelClassMap<DateTimeValue>
        {
            public CustomDateTimeClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(new DateTime(2017, 07, 20))
                    .WithInvalidFallback(new DateTime(2017, 07, 21));
            }
        }

        private class DateTimeFormatsArrayMap : ExcelClassMap<CustomDateTimeValue>
        {
            public DateTimeFormatsArrayMap()
            {
                Map(o => o.CustomValue)
                    .WithDateFormats("yyyy-MM-dd", "G");
            }
        }

        private class DateTimeEnumerableFormatsMap : ExcelClassMap<CustomDateTimeValue>
        {
            public DateTimeEnumerableFormatsMap()
            {
                Map(o => o.CustomValue)
                    .WithDateFormats(new List<string> { "yyyy-MM-dd", "G" });
            }
        }

        private class DefaultNullableDateTimeClassMap : ExcelClassMap<NullableDateTimeValue>
        {
            public DefaultNullableDateTimeClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomNullableDateTimeClassMap : ExcelClassMap<NullableDateTimeValue>
        {
            public CustomNullableDateTimeClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(new DateTime(2017, 07, 20))
                    .WithInvalidFallback(new DateTime(2017, 07, 21));
            }
        }
    }
}
