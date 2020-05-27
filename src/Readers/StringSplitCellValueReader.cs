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
    public class StringSplitCellValueReader : SplitCellValueReader
    {
        private string[] _separators = new string[] { "," };

        /// <summary>
        /// Gets or sets the separators used to split the string value of the cell.
        /// </summary>
        public string[] Separators
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
        /// Constructs a reader that reads the string value of a cell and produces multiple values
        /// by splitting it.
        /// </summary>
        /// <param name="cellReader">The ICellValueReader that reads the string value of the cell before it is split.</param>
        public StringSplitCellValueReader(ISingleCellValueReader cellReader) : base(cellReader)
        {
        }

        protected override string[] GetValues(string value) => value.Split(Separators, Options);
    }
}
