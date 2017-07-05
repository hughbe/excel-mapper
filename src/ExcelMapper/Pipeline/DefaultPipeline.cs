using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelMapper.Pipeline.Items;

namespace ExcelMapper.Pipeline
{
    public sealed class DefaultPipeline<TProperty> : ColumnPipeline<TProperty>
    {
        private SinglePipeline<TProperty> Override { get; set; }

        public DefaultPipeline(MemberInfo member) : base(member.Name, member)
        {
            AutoMapper.AutoMap(this);
        }

        public ColumnPipeline<TProperty> WithColumnName(string columnName)
        {
            var pipeline = new ColumnPipeline<TProperty>(columnName, Member)
            {
                Items = Items,
                InvalidFallback = InvalidFallback,
                EmptyFallback = EmptyFallback
            };
            Override = pipeline;

            return pipeline;
        }

        public IndexPipeline<TProperty> WithIndex(int index)
        {
            var pipeline = new IndexPipeline<TProperty>(index, Member)
            {
                Items = Items,
                InvalidFallback = InvalidFallback,
                EmptyFallback = EmptyFallback
            };
            Override = pipeline;

            return pipeline;
        }

        protected internal override object Execute(ExcelSheet sheet, ExcelRow row)
        {
            if (Override != null)
            {
                return Override.Execute(sheet, row);
            }

            return base.Execute(sheet, row);
        }
    }
}
