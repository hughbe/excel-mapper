using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Items;

namespace ExcelMapper.Utilities
{
    internal static class AutoMapper
    {
        public static void AutoMap(this SinglePropertyMapping pipeline, EmptyValueStrategy emptyValueStrategy)
        {
            // String nullable from types.
            Type type = pipeline.Type;
            bool isNullable = false;
            if (type.IsNullable())
            {
                isNullable = true;
                type = type.GenericTypeArguments[0];
            }

            Type[] interfaces = type.GetTypeInfo().ImplementedInterfaces.ToArray();

            EmptyValueStrategy emptyStrategyToPursue = EmptyValueStrategy.SetToDefaultValue;

            if (type == typeof(DateTime))
            {
                var item = new ParseAsDateTimeMappingItem();
                pipeline.AddMappingItem(item);

                if (!isNullable)
                {
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline = pipeline.WithThrowingInvalidFallback();
            }
            else if (type == typeof(bool))
            {
                var item = new ParseAsBoolMappingItem();
                pipeline.AddMappingItem(item);

                if (!isNullable)
                {
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline = pipeline.WithThrowingInvalidFallback();
            }
            else if (type.GetTypeInfo().IsEnum)
            {
                var item = new ParseAsEnumMappingItem(type);
                pipeline.AddMappingItem(item);

                if (!isNullable)
                {
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline = pipeline.WithThrowingInvalidFallback();
            }
            else if (type == typeof(string))
            {
                var item = new ParseAsStringMappingItem();
                pipeline.AddMappingItem(item);
            }
            else if (interfaces.Any(t => t == typeof(IConvertible)))
            {
                var item = new ChangeTypeMappingItem(type);
                pipeline.AddMappingItem(item);

                if (!isNullable)
                {
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline = pipeline.WithThrowingInvalidFallback();
            }
            else
            {
                throw new ExcelMappingException($"Don't know how to map type {type}.");
            }

            if (emptyStrategyToPursue == EmptyValueStrategy.SetToDefaultValue || emptyValueStrategy == EmptyValueStrategy.SetToDefaultValue)
            {
                var fallback = new FixedValueFallback(isNullable ? null : type.DefaultValue());
                pipeline.EmptyFallback = fallback;
            }
            else
            {
                Debug.Assert(emptyValueStrategy == EmptyValueStrategy.ThrowIfPrimitive);

                // The user specified that we should set to the default value if it was empty.
                pipeline = pipeline.WithThrowingEmptyFallback();
            }
        }
    }
}
