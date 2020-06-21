using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper
{
    public abstract class OneToOneMap : Map
    {
        public OneToOneMap(ISingleCellValueReader reader)
        {
            CellReader = reader ?? throw new ArgumentNullException(nameof(reader));
        }

        private ISingleCellValueReader _reader;

        public ISingleCellValueReader CellReader
        {
            get => _reader;
            set => _reader = value ?? throw new ArgumentNullException(nameof(value));
        }

        public bool Optional { get; set; }
    }
}
