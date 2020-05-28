using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks
{
    public class ReadRows
    {
        [Benchmark]
        public void DefaultMap()
        {
            using var original = Helpers.GetResource("VeryLargeSheet.xlsx");
            var importer = new ExcelImporter(original);
            importer.Configuration.RegisterClassMap<DataClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            foreach (object value in sheet.ReadRows<DataClass>())
            {
            }
        }

        [Benchmark]
        public void SkipBlankLinesMap()
        {
            using var original = Helpers.GetResource("VeryLargeSheet.xlsx");
            var importer = new ExcelImporter(original);
            importer.Configuration.SkipBlankLines = true;
            importer.Configuration.RegisterClassMap<DataClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            foreach (object value in sheet.ReadRows<DataClass>())
            {
            }
        }

        [Benchmark]
        public void OptionalMap()
        {
            using var original = Helpers.GetResource("VeryLargeSheet.xlsx");
            var importer = new ExcelImporter(original);
            importer.Configuration.RegisterClassMap<OptionalDataClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            foreach (object value in sheet.ReadRows<DataClass>())
            {
            }
        }

        [Benchmark]
        public void OptionalNoSuchValueMap()
        {
            using var original = Helpers.GetResource("VeryLargeSheet.xlsx");
            var importer = new ExcelImporter(original);
            importer.Configuration.RegisterClassMap<OptionalDataWithMissingValueClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            foreach (object value in sheet.ReadRows<DataClassWithMissingValue>())
            {
            }
        }

        private class DataClass
        {
            public int Value { get; set; }
        }

        private class DataClassWithMissingValue
        {
            public int Value { get; set; }
            public int NoSuchValue { get; set; }
        }

        private class DataClassMap : ExcelClassMap<DataClass>
        {
            public DataClassMap()
            {
                Map(p => p.Value);
            }
        }

        private class OptionalDataClassMap : ExcelClassMap<DataClass>
        {
            public OptionalDataClassMap()
            {
                Map(p => p.Value);
            }
        }

        private class OptionalDataWithMissingValueClassMap : ExcelClassMap<DataClassWithMissingValue>
        {
            public OptionalDataWithMissingValueClassMap()
            {
                Map(p => p.Value);

                Map(p => p.NoSuchValue)
                    .MakeOptional()
                    .WithEmptyFallback(0);
            }
        }
    }
}
