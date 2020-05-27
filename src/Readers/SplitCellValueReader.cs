using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers
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

        private ISingleCellValueReader _cellReader;

        /// <summary>
        /// Gets or sets the ICellValueReader that reads the string value of the cell
        /// before it is split.
        /// </summary>
        public ISingleCellValueReader CellReader
        {
            get => _cellReader;
            set => _cellReader = value ?? throw new ArgumentNullException(nameof(value));
        }

        /// <summary>
        /// Constructs a reader that reads the string value of a cell and produces multiple values
        /// by splitting it.
        /// </summary>
        /// <param name="cellReader">The ICellValueReader that reads the string value of the cell before it is split.</param>
        public SplitCellValueReader(ISingleCellValueReader cellReader)
        {
            CellReader = cellReader ?? throw new ArgumentNullException(nameof(cellReader));
        }

        public bool TryGetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, out IEnumerable<ReadCellValueResult> result)
        {
            if (!CellReader.TryGetValue(sheet, rowIndex, reader, out ReadCellValueResult readResult))
            {
                result = default;
                return false;
            }

            if (readResult.StringValue == null)
            {
                result = Enumerable.Empty<ReadCellValueResult>();
                return true;
            }

            string[] splitStringValues = GetValues(readResult.StringValue);
            result = splitStringValues.Select(s => new ReadCellValueResult(readResult.ColumnIndex, s));
            return true;
        }

        protected abstract string[] GetValues(string value);
    }
}
