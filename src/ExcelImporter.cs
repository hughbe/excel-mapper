using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;

namespace ExcelMapper
{
    /// <summary>
    /// An importer that reads the sheets in an Excel file or stream.
    /// </summary>
    public class ExcelImporter : IDisposable
    {
        /// <summary>
        /// Gets the inner reader that the importer wraps. This is defined by the ExcelDataReader library.
        /// </summary>
        public IExcelDataReader Reader { get; }

        /// <summary>
        /// Gets the configuration for the importer to allow customizing the importer. For example you
        /// can register custom class maps and specify whether sheets have a header.
        /// </summary>
        public ExcelImporterConfiguration Configuration { get; } = new ExcelImporterConfiguration();

        private int SheetIndex { get; set; } = -1;

        /// <summary>
        /// Constructs an importer that reads an Excel file from a stream.
        /// </summary>
        /// <param name="stream">A stream containing the Excel file bytes.</param>
        public ExcelImporter(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            Reader = ExcelReaderFactory.CreateReader(stream);
        }

        /// <summary>
        /// Constructs an importer that reads an Excel file from an existing data reader.
        /// </summary>
        /// <param name="reader">The existing data reader that wraps an Excel file.</param>
        public ExcelImporter(IExcelDataReader reader)
        {
            Reader = reader ?? throw new ArgumentNullException(nameof(reader));
        }

        /// <summary>
        /// Cleans up resources associated with this class. This primarily involves disposing of
        /// the inner reader.
        /// </summary>
        public void Dispose() => Reader.Dispose();

        /// <summary>
        /// Reads each sheet in the document.
        /// </summary>
        /// <returns>A lazily evaluated list of each sheet in the document.</returns>
        public IEnumerable<ExcelSheet> ReadSheets()
        {
            while (TryReadSheet(out ExcelSheet sheet))
            {
                yield return sheet;
            }
        }

        /// <summary>
        /// Reads the next sheet in the document. If no sheets have been read, then this reads the first sheet.
        /// The method throws if there are no more sheets in the document.
        /// </summary>
        /// <returns>The next sheet in the document.</returns>
        public ExcelSheet ReadSheet()
        {
            if (!TryReadSheet(out ExcelSheet sheet))
            {
                throw new ExcelMappingException("No more sheets.");
            }

            return sheet;
        }

        /// <summary>
        /// Reads the next sheet in the document. If no sheets have been read, then this reads the first sheet.
        /// </summary>
        /// <param name="excelSheet">The next sheet in the document.</param>
        /// <returns>False if there are no more sheets in the document, else true.</returns>
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
