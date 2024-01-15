using Xunit;

namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit {}
}

namespace ExcelMapper.Tests
{
    public class MapRecordTests
    {
        [Fact]
        public void ReadRow_AutoMappedRecord_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Numbers.xlsx");
            importer.Configuration.RegisterClassMap<DefaultRecordClassMap>();
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            var row1 = sheet.ReadRow<RecordClass>();
            Assert.Equal(2, row1.IntRecord.Value);
            
            // Empty value
            var row2 = sheet.ReadRow<RecordClass>();
            Assert.Null(row2.IntRecord);

            // Invalid value
            var row3 = sheet.ReadRow<RecordClass>();
            Assert.Null(row3.IntRecord);
        }

        private record IntRecord(int Value);

        private class RecordClass
        {
            public IntRecord IntRecord { get; private set; }
        }

        private class DefaultRecordClassMap : ExcelClassMap<RecordClass>
        {
            public DefaultRecordClassMap()
            {
                Map(u => u.IntRecord)
                    .WithConverter(v => new IntRecord(int.Parse(v)))
                    .WithColumnName("Value");
            }
        }
    }
}
