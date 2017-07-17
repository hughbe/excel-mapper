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
    public class SplitCellValueReader : IMultipleCellValuesReader
    {
        private char[] _separators = new char[] { ',' };

        /// <summary>
        /// Gets or sets the separators used to split the string value of the cell.
        /// </summary>
        public char[] Separators
        {
            get => _separators;
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }

                if (value.Length == 0)
                {
                    throw new ArgumentException("Separators cannot be empty.", nameof(value));
                }

                _separators = value;
            }
        }

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

            string[] splitStringValues = readResult.StringValue.Split(Separators, Options);
            return splitStringValues.Select(s => new ReadCellValueResult(readResult.ColumnIndex, s));
        }
    }
}
