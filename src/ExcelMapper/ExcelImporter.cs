using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;

namespace ExcelMapper
{
    public class ExcelImporter : IDisposable
    {
        public IExcelDataReader Reader { get; }
        public ExcelImporterConfiguration Configuration { get; } = new ExcelImporterConfiguration();
        private int SheetIndex { get; set; } = -1;

        public ExcelImporter(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            Reader = ExcelReaderFactory.CreateReader(stream);
        }

        public ExcelImporter(IExcelDataReader reader)
        {
            Reader = reader ?? throw new ArgumentNullException(nameof(reader));
        }

        public void Dispose() => Reader.Dispose();

        public IEnumerable<ExcelSheet> ReadSheets()
        {
            while (TryReadSheet(out ExcelSheet sheet))
            {
                yield return sheet;
            }
        }

        public ExcelSheet ReadSheet()
        {
            if (!TryReadSheet(out ExcelSheet sheet))
            {
                throw new ExcelMappingException("No more sheets.");
            }

            return sheet;
        }

        public bool TryReadSheet(out ExcelSheet excelSheet)
        {
            excelSheet = null;

            if (SheetIndex != -1)
            {
                if (!Reader.NextResult())
                {
                    return false;
                }
            }

            SheetIndex++;
            excelSheet = new ExcelSheet(Reader, SheetIndex, Configuration);
            return true;
        }

        private IEnumerable<T> ReadTraining<T>(IExcelDataReader reader)
        {
            foreach (Dictionary<string, object> result in ReadRows(reader))
            {
                Console.WriteLine(result.Count);
            }

            return null;
        }

        private static IEnumerable<Dictionary<string, object>> ReadRows(IExcelDataReader reader)
        {
            Dictionary<string, int> columnMappings = GetColumnMappings(reader);
            while (reader.Read())
            {
                yield return ReadRow(reader, columnMappings);
            }
        }

        private static Dictionary<string, object> ReadRow(IExcelDataReader reader, Dictionary<string, int> columnMappings)
        {
            var result = new Dictionary<string, object>(columnMappings.Count);
            foreach (KeyValuePair<string, int> mapping in columnMappings)
            {
                string value = reader.GetString(mapping.Value);
                result.Add(mapping.Key, value);
            }

            return result;
        }

        private static Dictionary<string, int> GetColumnMappings(IExcelDataReader reader)
        {
            reader.Read();

            var columnNames = new Dictionary<string, int>(reader.FieldCount);
            for (int columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
            {
                string name = reader.GetString(columnIndex);
                columnNames.Add(name, columnIndex);
            }

            return columnNames;
        }
    }
}
