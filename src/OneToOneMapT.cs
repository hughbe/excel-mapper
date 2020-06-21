using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper
{
    public class OneToOneMap<T> : OneToOneMap, IValuePipeline<T>
    {
        public OneToOneMap(ISingleCellValueReader reader) : base(reader)
        {
        }

        public ValuePipeline<T> Pipeline { get; } = new ValuePipeline<T>();

        public override bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo member, out object result)
        {
            if (!CellReader.TryGetValue(sheet, rowIndex, reader, out ReadCellValueResult readResult))
            {
                if (Optional)
                {
                    result = default;
                    return false;
                }

                throw new ExcelMappingException($"Could not read value for {member.Name}", sheet, rowIndex);
            }

            result = (T)ValuePipeline.GetPropertyValue(Pipeline, sheet, rowIndex, reader, readResult, member);
            return true;
        }

        public IEnumerable<ICellValueTransformer> CellValueTransformers => Pipeline.CellValueTransformers;

        public IEnumerable<ICellValueMapper> CellValueMappers => Pipeline.CellValueMappers;

        public IFallbackItem EmptyFallback
        {
            get => Pipeline.EmptyFallback;
            set => Pipeline.EmptyFallback = value;
        }

        public IFallbackItem InvalidFallback
        {
            get => Pipeline.InvalidFallback;
            set => Pipeline.InvalidFallback = value;
        }

        public void AddCellValueMapper(ICellValueMapper mapper) => Pipeline.AddCellValueMapper(mapper);

        public void AddCellValueTransformer(ICellValueTransformer transformer) => Pipeline.AddCellValueTransformer(transformer);

        public void RemoveCellValueMapper(int index) => Pipeline.RemoveCellValueMapper(index);
    }
}
