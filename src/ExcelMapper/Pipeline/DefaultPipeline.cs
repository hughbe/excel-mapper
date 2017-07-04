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
            AutoMap();
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

        private void AutoMap()
        {
            var pipelineItems = new List<PipelineItem<TProperty>>();

            void AddItem(PipelineItem<TProperty> item)
            {
                item.Automapped = true;
                pipelineItems.Add(item);
            }

            Type type = typeof(TProperty);
            Type[] interfaces = type.GetTypeInfo().ImplementedInterfaces.ToArray();

            if (type == typeof(DateTime))
            {
                var item = new ParseAsDateTimePipelineItem() as PipelineItem<TProperty>;
                AddItem(item);

                this.DisallowEmptyOrInvalid<DefaultPipeline<TProperty>, TProperty>();
            }
            else if (type == typeof(bool))
            {
                var item = new ParseAsBoolPipelineItem() as PipelineItem<TProperty>;
                AddItem(item);

                this.DisallowEmptyOrInvalid<DefaultPipeline<TProperty>, TProperty>();
            }
            else if (type.GetTypeInfo().BaseType == typeof(Enum))
            {
                Type itemType = typeof(ParseAsEnumPipelineItem<>).MakeGenericType(type);
                object item = Activator.CreateInstance(itemType);
                AddItem((PipelineItem<TProperty>)item);

                this.DisallowEmptyOrInvalid<DefaultPipeline<TProperty>, TProperty>();
            }
            else if (type == typeof(string))
            {
                AddItem(ParseAsStringPipelineItem.Instance as PipelineItem<TProperty>);
            }
            else if (interfaces.Any(t => t == typeof(IConvertible)))
            {
                var item = new ChangeTypePipelineItem<TProperty>();
                AddItem(item);

                this.DisallowEmptyOrInvalid<DefaultPipeline<TProperty>, TProperty>();
            }
            else
            {
                Type ienumerableType = interfaces.FirstOrDefault(t => t.GetTypeInfo().IsGenericType && t.GetGenericTypeDefinition() == typeof(IEnumerable<>));
                if (ienumerableType != null)
                {
                    Type elementType = ienumerableType.GenericTypeArguments[0];
                    Type parserType = typeof(SplitWithDelimiterPipelineItem<,>).MakeGenericType(typeof(TProperty), elementType);

                    PipelineItem<TProperty> parser = (PipelineItem<TProperty>)Activator.CreateInstance(parserType);
                    AddItem(parser);
                }
            }

            this.WithAdditionalItems(pipelineItems);
        }
    }
}
