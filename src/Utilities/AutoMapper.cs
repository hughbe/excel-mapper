using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Mappers;
using ExcelMapper.Mappings.MultiItems;

namespace ExcelMapper.Utilities
{
    internal static class AutoMapper
    {
        private static MethodInfo s_mappingMethod;
        private static MethodInfo MappingMethod => s_mappingMethod ?? (s_mappingMethod = typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(InferMapping)));

        private static MethodInfo s_tryMapEnumerableMethod;
        private static MethodInfo TryMapEnumerableMethod => s_tryMapEnumerableMethod ?? (s_tryMapEnumerableMethod = typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(TryMapGenericEnumerable)));

        private static bool InferMapping<T>(MemberInfo member, FallbackStrategy emptyValueStrategy, out ExcelPropertyMap map)
        {
            // First, check if this is a well-known type (e.g. string/int).
            // This is a simple conversion from the cell's value to the type.
            if (TryMapPrimitive(member, emptyValueStrategy, out SingleExcelPropertyMap<T> singleMap))
            {
                map = singleMap;
                return true;
            }

            // Secondly, check if this is a collection (e.g. array, list).
            // This requires converting each value to the element type of the collection.
            if (TryMapEnumerable(member, emptyValueStrategy, out ExcelPropertyMap multiMap))
            {
                map = multiMap;
                return true;
            }

            // Thirdly, check if this is an object.
            // This requires converting each member and setting it on the object.
            if (TryMapObject(member, emptyValueStrategy, out ObjectExcelPropertyMap<T> objectMap))
            {
                map = objectMap;
                return true;
            }

            map = null;
            return false;
        }

        public static bool TryMapPrimitive<T>(MemberInfo member, FallbackStrategy emptyValueStrategy, out SingleExcelPropertyMap<T> map)
        {
            if (!GetWellKnownMapper(typeof(T), emptyValueStrategy, out ICellValueMapper mapper, out IFallbackItem emptyFallback, out IFallbackItem invalidFallback))
            {
                map = null;
                return false;
            }

            map = new SingleExcelPropertyMap<T>(member)
                .WithCellValueMappers(mapper)
                .WithEmptyFallbackItem(emptyFallback)
                .WithInvalidFallbackItem(invalidFallback);
            return true;
        }

        private static bool TryMapEnumerable(MemberInfo member, FallbackStrategy emptyValueStrategy, out ExcelPropertyMap map)
        {
            if (!member.MemberType().GetElementTypeOrEnumerableType(out Type elementType))
            {
                map = null;
                return false;
            }

            MethodInfo method = TryMapEnumerableMethod.MakeGenericMethod(elementType);

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

        private static bool GetWellKnownMapper(Type memberType, FallbackStrategy emptyValueStrategy, out ICellValueMapper mapper, out IFallbackItem emptyFallback, out IFallbackItem invalidFallback)
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

        public static bool TryMapGenericEnumerable<T>(MemberInfo member, FallbackStrategy emptyValueStrategy, out EnumerableExcelPropertyMap<T> map)
        {
            Type rawType = member.MemberType();
            TypeInfo rawTypeInfo = rawType.GetTypeInfo();

            // First, get the mapper for the element. This is used to convert individual values
            // to be added to/included in the collection.
            if (!TryMapPrimitive(member, emptyValueStrategy, out SingleExcelPropertyMap<T> elementMapping))
            {
                map = null;
                return false;
            }

            // Secondly, find the right way of adding the converted value to the collection.
            if (rawType.IsArray)
            {
                // Add values using the arrray indexer.
                map = new ArrayPropertyMap<T>(member, elementMapping);
                return true;
            }
            else if (rawTypeInfo.IsInterface)
            {
                // Add values by creating a list and assigning to the property.
                if (rawTypeInfo.IsAssignableFrom(typeof(List<T>).GetTypeInfo()))
                {
                    map = new InterfaceAssignableFromListPropertyMap<T>(member, elementMapping);
                    return true;
                }
            }
            else if (rawType.ImplementsInterface(typeof(ICollection<T>)))
            {
                // Add values using the ICollection<T>.Add method.
                map = new ConcreteICollectionPropertyMap<T>(rawType, member, elementMapping);
                return true;
            }

            map = null;
            return false;
        }

        public static bool TryMapObject<T>(MemberInfo member, FallbackStrategy emptyValueStrategy, out ObjectExcelPropertyMap<T> mapping)
        {
            if (!GenerateObjectMap(emptyValueStrategy, out ExcelClassMap<T> excelClassMap))
            {
                mapping = null;
                return false;
            }

            mapping = new ObjectExcelPropertyMap<T>(member, excelClassMap);
            return true;
        }

        public static bool GenerateObjectMap<T>(FallbackStrategy emptyValueStrategy, out ExcelClassMap<T> classMap)
        {
            Type type = typeof(T);

            if (type.GetTypeInfo().IsInterface)
            {
                classMap = null;
                return false;
            }

            var map = new ExcelClassMap<T>();
            IEnumerable<MemberInfo> properties = type.GetRuntimeProperties().Where(p => p.CanWrite);
            IEnumerable<MemberInfo> fields = type.GetRuntimeFields().Where(f => f.IsPublic);

            foreach (MemberInfo member in properties.Concat(fields))
            {
                // Infer the mapping for each member (property/field) belonging to the type.
                Type memberType = member.MemberType();
                MethodInfo method = MappingMethod.MakeGenericMethod(memberType);

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
