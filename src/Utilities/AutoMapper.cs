using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;

namespace ExcelMapper.Utilities
{
    public static class AutoMapper
    {
        private static MethodInfo s_tryCreateMemberMapMethod;
        private static MethodInfo TryCreateMemberMapMethod => s_tryCreateMemberMapMethod ?? (s_tryCreateMemberMapMethod = typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(TryCreateMemberMap)));

        private static MethodInfo s_tryCreateGenericEnumerableMapMethod;
        private static MethodInfo TryCreateGenericEnumerableMapMethod => s_tryCreateGenericEnumerableMapMethod ?? (s_tryCreateGenericEnumerableMapMethod = typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(TryCreateGenericEnumerableMap)));

        private static bool TryCreateMemberMap<T>(MemberInfo member, FallbackStrategy emptyValueStrategy, out ExcelPropertyMap map)
        {
            // First, check if this is a well-known type (e.g. string/int).
            // This is a simple conversion from the cell's value to the type.
            if (TryCreatePrimitiveMap(member, emptyValueStrategy, out OneToOnePropertyMap<T> singleMap))
            {
                map = singleMap;
                return true;
            }

            // Secondly, check if this is a collection (e.g. array, list).
            // This requires converting each value to the element type of the collection.
            if (TryCreateEnumerableMap(member, emptyValueStrategy, out ExcelPropertyMap multiMap))
            {
                map = multiMap;
                return true;
            }

            // Thirdly, check if this is an object.
            // This requires converting each member and setting it on the object.
            if (TryCreateObjectMap(member, emptyValueStrategy, out ManyToOneObjectPropertyMap<T> objectMap))
            {
                map = objectMap;
                return true;
            }

            map = null;
            return false;
        }

        internal static bool TryCreatePrimitivePipeline<T>(FallbackStrategy emptyValueStrategy, out ValuePipeline<T> pipeline)
        {
            if (!TryGetWellKnownMap(typeof(T), emptyValueStrategy, out ICellValueMapper mapper, out IFallbackItem emptyFallback, out IFallbackItem invalidFallback))
            {
                pipeline = null;
                return false;
            }

            pipeline = new ValuePipeline<T>();
            pipeline.AddCellValueMapper(mapper);
            pipeline.EmptyFallback = emptyFallback;
            pipeline.InvalidFallback = invalidFallback;
            return true;
        }

        internal static bool TryCreatePrimitiveMap<T>(MemberInfo member, FallbackStrategy emptyValueStrategy, out OneToOnePropertyMap<T> map)
        {
            if (!TryGetWellKnownMap(typeof(T), emptyValueStrategy, out ICellValueMapper mapper, out IFallbackItem emptyFallback, out IFallbackItem invalidFallback))
            {
                map = null;
                return false;
            }

            map = new OneToOnePropertyMap<T>(member)
                .WithCellValueMappers(mapper)
                .WithEmptyFallbackItem(emptyFallback)
                .WithInvalidFallbackItem(invalidFallback);
            return true;
        }

        private static bool TryCreateEnumerableMap(MemberInfo member, FallbackStrategy emptyValueStrategy, out ExcelPropertyMap map)
        {
            if (!member.MemberType().GetElementTypeOrEnumerableType(out Type elementType))
            {
                map = null;
                return false;
            }

            MethodInfo method = TryCreateGenericEnumerableMapMethod.MakeGenericMethod(elementType);

            var parameters = new object[] { member, emptyValueStrategy, null };
            bool result = (bool)method.Invoke(null, parameters);
            if (result)
            {
                map = (ExcelPropertyMap)parameters[2];
                return true;
            }

            map = null;
            return false;
        }

        private static bool TryGetWellKnownMap(Type memberType, FallbackStrategy emptyValueStrategy, out ICellValueMapper mapper, out IFallbackItem emptyFallback, out IFallbackItem invalidFallback)
        {
            Type type = memberType.GetNullableTypeOrThis(out bool isNullable);
            Type[] interfaces = type.GetTypeInfo().ImplementedInterfaces.ToArray();

            IFallbackItem ReconcileFallback(FallbackStrategy strategyToPursue, bool isEmpty)
            {
                // Empty nullable values should be set to null.
                if (isEmpty && isNullable)
                {
                    return new FixedValueFallback(null);
                }
                else if (strategyToPursue == FallbackStrategy.SetToDefaultValue || emptyValueStrategy == FallbackStrategy.SetToDefaultValue)
                {
                    return new FixedValueFallback(type.DefaultValue());
                }
                else
                {
                    Debug.Assert(emptyValueStrategy == FallbackStrategy.ThrowIfPrimitive);

                    // The user specified that we should set to the default value if it was empty.
                    return new ThrowFallback();
                }
            }

            // Set the default mapper for each well-known type.
            if (type == typeof(DateTime))
            {
                mapper = new DateTimeMapper();
                emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
            }
            else if (type == typeof(Guid))
            {
                mapper = new GuidMapper();
                emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
            }
            else if (type == typeof(bool))
            {
                mapper = new BoolMapper();
                emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
            }
            else if (type.GetTypeInfo().IsEnum)
            {
                mapper = new EnumMapper(type);
                emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
            }
            else if (type == typeof(string) || type == typeof(object) || type == typeof(IConvertible))
            {
                mapper = new StringMapper();
                emptyFallback = ReconcileFallback(FallbackStrategy.SetToDefaultValue, isEmpty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.SetToDefaultValue, isEmpty: false);
            }
            else if (type == typeof(Uri))
            {
                mapper = new UriMapper();
                emptyFallback = ReconcileFallback(FallbackStrategy.SetToDefaultValue, isEmpty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
            }
            else if (interfaces.Any(t => t == typeof(IConvertible)))
            {
                mapper = new ChangeTypeMapper(type);
                emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, isEmpty: false);
            }
            else
            {
                mapper = null;
                emptyFallback = null;
                invalidFallback = null;
                return false;
            }

            return true;
        }

        internal static bool TryCreateGenericEnumerableMap<TElement>(MemberInfo member, FallbackStrategy emptyValueStrategy, out ManyToOneEnumerablePropertyMap<TElement> map)
        {
            Type rawType = member.MemberType();

            // First, get the pipeline for the element. This is used to convert individual values
            // to be added to/included in the collection.
            if (!TryCreatePrimitivePipeline<TElement>(emptyValueStrategy, out ValuePipeline<TElement> elementMapping))
            {
                map = null;
                return false;
            }

            // Secondly, find the right way of adding the converted value to the collection.
            if (!TryGetCreateElementsFactory<TElement>(rawType, out CreateElementsFactory<TElement> factory))
            {
                map = null;
                return false;
            }

            // Default to splitting.
            var defaultNameReader = new ColumnNameValueReader(member.Name);
            var defaultReader = new CharSplitCellValueReader(defaultNameReader);
            map = new ManyToOneEnumerablePropertyMap<TElement>(member, defaultReader, elementMapping, factory);
            return true;
        }

        private static bool TryGetCreateElementsFactory<T>(Type memberType, out CreateElementsFactory<T> result)
        {
            if (memberType.IsArray)
            {
                result = elements => elements.ToArray();
                return true;
            }
            else if (memberType.GetTypeInfo().IsInterface)
            {
                // Add values by creating a list and assigning to the property.
                if (memberType.GetTypeInfo().IsAssignableFrom(typeof(List<T>).GetTypeInfo()))
                {
                    result = elements => elements;
                    return true;
                }
            }
            else if (memberType.ImplementsInterface(typeof(ICollection<T>)))
            {
                result = elements =>
                {
                    ICollection<T> value = (ICollection<T>)Activator.CreateInstance(memberType);
                    foreach (T element in elements)
                    {
                        value.Add(element);
                    }

                    return value;
                };
                return true;
            }

            result = default;
            return false;
        }

        internal static bool TryCreateObjectMap<T>(MemberInfo member, FallbackStrategy emptyValueStrategy, out ManyToOneObjectPropertyMap<T> mapping)
        {
            if (!TryCreateClassMap(emptyValueStrategy, out ExcelClassMap<T> excelClassMap))
            {
                mapping = null;
                return false;
            }

            mapping = new ManyToOneObjectPropertyMap<T>(member, excelClassMap);
            return true;
        }

        /// <summary>
        /// Creates a class map for the given type using the given strategy.
        /// </summary>
        /// <param name="emptyValueStrategy">The default strategy to use when the value of a cell is empty.</param>
        /// <param name="classMap">The class map for the given type.</param>
        /// <returns>True if the class map could be created, else false.</returns>
        public static bool TryCreateClassMap<T>(FallbackStrategy emptyValueStrategy, out ExcelClassMap<T> classMap)
        {
            if (!Enum.IsDefined(typeof(FallbackStrategy), emptyValueStrategy))
            {
                throw new ArgumentException($"Invalid value \"{emptyValueStrategy}\".", nameof(emptyValueStrategy));
            }

            Type type = typeof(T);

            if (type.GetTypeInfo().IsInterface)
            {
                classMap = null;
                return false;
            }

            var map = new ExcelClassMap<T>(emptyValueStrategy);
            IEnumerable<MemberInfo> properties = type.GetRuntimeProperties().Where(p => p.CanWrite && p.SetMethod.IsPublic && !p.SetMethod.IsStatic);
            IEnumerable<MemberInfo> fields = type.GetRuntimeFields().Where(f => f.IsPublic && !f.IsStatic);

            foreach (MemberInfo member in properties.Concat(fields))
            {
                // Infer the mapping for each member (property/field) belonging to the type.
                Type memberType = member.MemberType();
                MethodInfo method = TryCreateMemberMapMethod.MakeGenericMethod(memberType);

                var parameters = new object[] { member, emptyValueStrategy, null };
                bool result = (bool)method.Invoke(null, parameters);
                if (!result)
                {
                    classMap = null;
                    return false;
                }

                // Get the out parameter representing the property map for the member.
                map.Mappings.Add((ExcelPropertyMap)parameters[2]);
            }

            classMap = map;
            return true;
        }
    }
}
