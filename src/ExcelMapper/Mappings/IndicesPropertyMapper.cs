using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Utilities;

namespace ExcelMapper.Mappings
{
    internal sealed class IndicesPropertyMapper : IMultiPropertyMapper
    {
        public int[] Indices { get; }

        internal IndicesPropertyMapper(IEnumerable<int> indices)
        {
            if (indices == null)
            {
                throw new ArgumentNullException(nameof(indices));
            }

            foreach (int columnIndex in indices)
            {
                if (columnIndex < 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(indices), columnIndex, $"Negative column index in {indices.ArrayJoin()}.");
                }
            }

            Indices = indices.ToArray();
        }

        public IEnumerable<MapResult> GetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            return Indices.Select(i => new MapResult(i, reader.GetString(i)));
        }
    }
}
