using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Mappers;

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
                var mapper = new DateTimeMapper();
                pipeline.AddMappingItem(mapper);

                if (!isNullable)
                {
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline = pipeline.WithThrowingInvalidFallback();
            }
            else if (type == typeof(bool))
            {
                var mapper = new BoolMapper();
                pipeline.AddMappingItem(mapper);

                if (!isNullable)
                {
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline = pipeline.WithThrowingInvalidFallback();
            }
            else if (type.GetTypeInfo().IsEnum)
            {
                var mapper = new EnumMapper(type);
                pipeline.AddMappingItem(mapper);

                if (!isNullable)
                {
                    emptyStrategyToPursue = EmptyValueStrategy.ThrowIfPrimitive;
                }

                pipeline = pipeline.WithThrowingInvalidFallback();
            }
            else if (type == typeof(string))
            {
                var mapper = new StringMapper();
                pipeline.AddMappingItem(mapper);
            }
            else if (interfaces.Any(t => t == typeof(IConvertible)))
            {
                var mapper = new ChangeTypeMapper(type);
                pipeline.AddMappingItem(mapper);

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
                var emptyFallback = new FixedValueFallback(isNullable ? null : type.DefaultValue());
                pipeline.EmptyFallback = emptyFallback;
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
