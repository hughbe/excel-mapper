using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Readers
{
    /// <summary>
    /// Reads the value of a cell and produces multiple values by splitting the string value
    /// using the given separators.
    /// </summary>
    public abstract class SplitCellValueReader : IMultipleCellValuesReader
    {
        /// <summary>
        /// Gets or sets the options used to split the string value of the cell.
        /// </summary>
        public StringSplitOptions Options { get; set; }

        private ICellValueReader _cellReader;

        /// <summary>
        /// Gets or sets the ICellValueReader that reads the string value of the cell
        /// before it is split.
        /// </summary>
        public ICellValueReader CellReader
        {
            get => _cellReader;
            set => _cellReader = value ?? throw new ArgumentNullException(nameof(value));
        }

        /// <summary>
        /// Constructs a reader that reads the string value of a cell and produces multiple values
        /// by splitting it.
        /// </summary>
        /// <param name="cellReader">The ICellValueReader that reads the string value of the cell before it is split.</param>
        public SplitCellValueReader(ICellValueReader cellReader)
        {
            CellReader = cellReader ?? throw new ArgumentNullException(nameof(cellReader));
        }

        public IEnumerable<ReadCellValueResult> GetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            ReadCellValueResult readResult = CellReader.GetValue(sheet, rowIndex, reader);
            if (readResult.StringValue == null)
            {
                return Enumerable.Empty<ReadCellValueResult>();
            }

            string[] splitStringValues = GetValues(readResult.StringValue);
            return splitStringValues.Select(s => new ReadCellValueResult(readResult.ColumnIndex, s));
        }

        protected abstract string[] GetValues(string value);
    }
}
