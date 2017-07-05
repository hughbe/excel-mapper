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
        public static void AutoMap<TProperty>(Pipeline<TProperty> pipeline, EmptyValueStrategy emptyValueStrategy)
        {
            var pipelineItems = new List<PipelineItem<TProperty>>();

            void AddItem(PipelineItem<TProperty> item)
            {
                item.Automapped = true;
                pipelineItems.Add(item);
            }

            Type type = typeof(TProperty);

            // String nullable from types.
            bool isNullable = false;
            if (type.GetTypeInfo().IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                isNullable = true;
                type = type.GenericTypeArguments[0];
            }

            Type[] interfaces = type.GetTypeInfo().ImplementedInterfaces.ToArray();

            EmptyValueStrategy emptyStrategyToPursue = EmptyValueStrategy.SetToDefaultValue;

            if (type == typeof(DateTime))
            {
                var item = new ParseAsDateTimePipelineItem();
                if (isNullable)
                {
                    AddItem(new CastPipelineItem<TProperty, DateTime>(item));
                }
                else
                {
                    AddItem(item as PipelineItem<TProperty>);
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline.WithThrowingInvalidFallback<Pipeline<TProperty>, TProperty>();
            }
            else if (type == typeof(bool))
            {
                var item = new ParseAsBoolPipelineItem();
                if (isNullable)
                {
                    AddItem(new CastPipelineItem<TProperty, bool>(item));
                }
                else
                {
                    AddItem(item as PipelineItem<TProperty>);
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline.WithThrowingInvalidFallback<Pipeline<TProperty>, TProperty>();
            }
            else if (type.GetTypeInfo().BaseType == typeof(Enum))
            {
                Type itemType = typeof(ParseAsEnumPipelineItem<>).MakeGenericType(type);
                object item = Activator.CreateInstance(itemType);

                if (isNullable)
                {
                    Type castType = typeof(CastPipelineItem<,>).MakeGenericType(typeof(TProperty), type);
                    object castItem = Activator.CreateInstance(castType, new object[] { item });
                    AddItem(castItem as PipelineItem<TProperty>);
                }
                else
                {
                    AddItem(item as PipelineItem<TProperty>);
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline.WithThrowingInvalidFallback<Pipeline<TProperty>, TProperty>();
            }
            else if (type == typeof(string))
            {
                AddItem(ParseAsStringPipelineItem.Instance as PipelineItem<TProperty>);
            }
            else if (interfaces.Any(t => t == typeof(IConvertible)))
            {
                var item = new ChangeTypePipelineItem<TProperty>(type);
                AddItem(item);

                if (isNullable)
                {
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline.WithThrowingInvalidFallback<Pipeline<TProperty>, TProperty>();
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

            if (emptyStrategyToPursue == EmptyValueStrategy.ThrowIfPrimitive)
            {
                // The user specified that we should set to the default value if it was empty.
                if (emptyValueStrategy == EmptyValueStrategy.SetToDefaultValue)
                {
                    pipeline.WithEmptyFallback(default(TProperty));
                }
                else
                {
                    pipeline.WithThrowingEmptyFallback<Pipeline<TProperty>, TProperty>();
                }
            }

            pipeline.WithAdditionalItems(pipelineItems);
        }
    }
}
