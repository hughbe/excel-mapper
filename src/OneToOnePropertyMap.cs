using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Readers;

namespace ExcelMapper
{
    /// <summary>
    /// Reads a single cell of an excel sheet and maps the value of the cell to the
    /// type of the property or field.
    /// </summary>
    public class OneToOnePropertyMap : ExcelPropertyMap
    {
        private ISingleCellValueReader _reader;

        public ValuePipeline Pipeline { get; } = new ValuePipeline();

        /// <summary>
        /// Gets or sets the object that takes a sheet and row index and produces the value of a cell.
        /// </summary>
        public ISingleCellValueReader CellReader
        {
            get => _reader;
            set => _reader = value ?? throw new ArgumentNullException(nameof(value));
        }

        /// <summary>
        /// Gets or sets whether mapping should fail silently and continue if the cell value cannot be
        /// found.
        /// </summary>
        public bool Optional { get; set; }

        /// <summary>
        /// Constructs a map that reads the value of a single cell and maps the value of the cell
        /// to the type of the property or field.
        /// </summary>
        /// <param name="member">The property or field to map the value of a single cell to.</param>
        public OneToOnePropertyMap(MemberInfo member, ValuePipeline pipeline = null) : base(member)
        {
            CellReader = new ColumnNameValueReader(member.Name);
            Pipeline = pipeline ?? new ValuePipeline();
        }

        public override void SetPropertyValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, object instance)
        {
            if (!CellReader.TryGetValue(sheet, rowIndex, reader, out ReadCellValueResult readResult))
            {
                if (Optional)
                {
                    return;
                }

                throw new ExcelMappingException($"Could not read value for {Member.Name}", sheet, rowIndex);
            }

            object result = Pipeline.GetPropertyValue(sheet, rowIndex, reader, readResult, Member);
            SetPropertyFactory(instance, result);
        }
    }
}
