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
    }
}
