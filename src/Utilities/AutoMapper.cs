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

        private static MethodInfo s_autoMapEnumerableMethod;
        private static MethodInfo AutoMapEnumerableMethod => s_autoMapEnumerableMethod ?? (s_autoMapEnumerableMethod = typeof(AutoMapper).GetTypeInfo().GetDeclaredMethod(nameof(AutoMapEnumerable)));

        private static bool InferMapping<T>(this MemberInfo member, FallbackStrategy emptyValueStrategy, out ExcelPropertyMap mapping)
        {
            if (member.AutoMap(emptyValueStrategy, out SingleExcelPropertyMap<T> singleMapping))
            {
                mapping = singleMapping;
                return true;
            }

            if (member.MemberType().GetElementTypeOrEnumerableType(out Type elementType))
            {
                MethodInfo method = AutoMapEnumerableMethod.MakeGenericMethod(elementType);

                var parameters = new object[] { member, emptyValueStrategy, null };
                bool result = (bool)method.Invoke(null, parameters);
                if (result)
                {
                    mapping = (ExcelPropertyMap)parameters[2];
                    return true;
                }
            }

            if (member.AutoMapObject(emptyValueStrategy, out ObjectExcelPropertyMap<T> objectMapping))
            {
                mapping = objectMapping;
                return true;
            }

            mapping = null;
            return false;
        }

        public static bool AutoMap<T>(this MemberInfo member, FallbackStrategy emptyValueStrategy, out SingleExcelPropertyMap<T> mapping)
        {
            if (!typeof(T).AutoMap(emptyValueStrategy, out ICellValueMapper mapper, out IFallbackItem emptyFallback, out IFallbackItem invalidFallback))
            {
                mapping = null;
                return false;
            }

            mapping = new SingleExcelPropertyMap<T>(member)
                .WithCellValueMappers(mapper)
                .WithEmptyFallbackItem(emptyFallback)
                .WithInvalidFallbackItem(invalidFallback);
            return true;
        }

        private static bool AutoMap(this Type memberType, FallbackStrategy emptyValueStrategy, out ICellValueMapper mapper, out IFallbackItem emptyFallback, out IFallbackItem invalidFallback)
        {
            Type type = memberType.GetNullableTypeOrThis(out bool isNullable);

            Type[] interfaces = type.GetTypeInfo().ImplementedInterfaces.ToArray();

            IFallbackItem ReconcileFallback(FallbackStrategy strategyToPursue, bool empty)
            {
                // Empty nullable values should be set to null.
                if (empty && isNullable)
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

            if (type == typeof(DateTime))
            {
                mapper = new DateTimeMapper();
                emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, empty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, empty: false);
            }
            else if (type == typeof(bool))
            {
                mapper = new BoolMapper();
                emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, empty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, empty: false);
            }
            else if (type.GetTypeInfo().IsEnum)
            {
                mapper = new EnumMapper(type);
                emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, empty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, empty: false);
            }
            else if (type == typeof(string) || type == typeof(object) || type == typeof(IConvertible))
            {
                mapper = new StringMapper();
                emptyFallback = ReconcileFallback(FallbackStrategy.SetToDefaultValue, empty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.SetToDefaultValue, empty: false);
            }
            else if (type == typeof(Uri))
            {
                mapper = new UriMapper();
                emptyFallback = ReconcileFallback(FallbackStrategy.SetToDefaultValue, empty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, empty: false);
            }
            else if (interfaces.Any(t => t == typeof(IConvertible)))
            {
                mapper = new ChangeTypeMapper(type);
                emptyFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, empty: true);
                invalidFallback = ReconcileFallback(FallbackStrategy.ThrowIfPrimitive, empty: false);
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

        public static bool AutoMapEnumerable<T>(this MemberInfo member, FallbackStrategy emptyValueStrategy, out EnumerableExcelPropertyMap<T> map)
        {
            Type rawType = member.MemberType();
            TypeInfo rawTypeInfo = rawType.GetTypeInfo();

            if (!member.AutoMap(emptyValueStrategy, out SingleExcelPropertyMap<T> elementMapping))
            {
                map = null;
                return false;
            }

            if (rawType.IsArray)
            {
                map = new ArrayPropertyMap<T>(member, elementMapping);
                return true;
            }
            else if (rawTypeInfo.IsInterface)
            {
                if (rawTypeInfo.IsAssignableFrom(typeof(List<T>).GetTypeInfo()))
                {
                    map = new InterfaceAssignableFromListPropertyMap<T>(member, elementMapping);
                    return true;
                }
            }
            else if (rawType.ImplementsInterface(typeof(ICollection<T>)))
            {
                map = new ConcreteICollectionPropertyMap<T>(rawType, member, elementMapping);
                return true;
            }

            map = null;
            return false;
        }

        public static bool AutoMapObject<T>(this MemberInfo member, FallbackStrategy emptyValueStrategy, out ObjectExcelPropertyMap<T> mapping)
        {
            if (!AutoMapClass(emptyValueStrategy, out ExcelClassMap<T> excelClassMap))
            {
                mapping = null;
                return false;
            }

            mapping = new ObjectExcelPropertyMap<T>(member, excelClassMap);
            return true;
        }

        public static bool AutoMapClass<T>(FallbackStrategy emptyValueStrategy, out ExcelClassMap<T> classMap)
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
                Type memberType = member.MemberType();
                MethodInfo method = MappingMethod.MakeGenericMethod(memberType);

                var parameters = new object[] { member, emptyValueStrategy, null };
                bool result = (bool)method.Invoke(null, parameters);
                if (!result)
                {
                    classMap = null;
                    return false;
                }

                map.Mappings.Add((ExcelPropertyMap)parameters[2]);
            }

            classMap = map;
            return true;
        }
    }
}
