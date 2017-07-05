using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelMapper.Pipeline.Items;
using ExcelMapper.Utilities;

namespace ExcelMapper.Pipeline
{
    internal static class AutoMapper
    {
        public static void AutoMap<TProperty>(Pipeline<TProperty> pipeline)
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

                pipeline.WithThrowingFallback<Pipeline<TProperty>, TProperty>();
            }
            else if (type == typeof(bool))
            {
                var item = new ParseAsBoolPipelineItem() as PipelineItem<TProperty>;
                AddItem(item);

                pipeline.WithThrowingFallback<Pipeline<TProperty>, TProperty>();
            }
            else if (type.GetTypeInfo().BaseType == typeof(Enum))
            {
                Type itemType = typeof(ParseAsEnumPipelineItem<>).MakeGenericType(type);
                object item = Activator.CreateInstance(itemType);
                AddItem((PipelineItem<TProperty>)item);

                pipeline.WithThrowingFallback<Pipeline<TProperty>, TProperty>();
            }
            else if (type == typeof(string))
            {
                AddItem(ParseAsStringPipelineItem.Instance as PipelineItem<TProperty>);
            }
            else if (interfaces.Any(t => t == typeof(IConvertible)))
            {
                var item = new ChangeTypePipelineItem<TProperty>();
                AddItem(item);

                pipeline.WithThrowingFallback<Pipeline<TProperty>, TProperty>();
            }
            else
            {
                Type elementType = interfaces.GetIEnumerableElementType();
                if (elementType != null)
                {
                    Type parserType = typeof(SplitWithDelimiterPipelineItem<,>).MakeGenericType(typeof(TProperty), elementType);

                    PipelineItem<TProperty> parser = (PipelineItem<TProperty>)Activator.CreateInstance(parserType);
                    AddItem(parser);
                }
            }

            pipeline.WithAdditionalItems<Pipeline<TProperty>, TProperty>(pipelineItems);
        }
    }
}
