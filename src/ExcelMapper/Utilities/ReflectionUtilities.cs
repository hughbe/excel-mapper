using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Utilities
{
    internal static class ReflectionUtilities
    {
        public static Type GetIEnumerableType(this Type type)
        {
            IEnumerable<Type> interfaces = type.GetTypeInfo().ImplementedInterfaces;
            Type ienumerableType = interfaces.FirstOrDefault(parameterType => parameterType.ImplementsGenericInterface(typeof(IEnumerable<>)));

            return ienumerableType?.GenericTypeArguments[0];
        }

        public static bool ImplementsInterface(this Type type, Type interfaceType)
        {
            return type.GetTypeInfo().ImplementedInterfaces.Any(t => t == interfaceType);
        }

        public static bool ImplementsGenericInterface(this Type type, Type interfaceType)
        {
            if (!type.GetTypeInfo().IsGenericType)
            {
                return false;
            }

            return type.GetGenericTypeDefinition() == interfaceType;
        }

        public static ConstructorInfo GetConstructor(this Type type, Type[] parameters)
        {
            IEnumerable<ConstructorInfo> constructors = type.GetTypeInfo().DeclaredConstructors;
            return constructors.First(c => c.HasParameters(parameters));
        }

        private static bool HasParameters(this MethodBase method, Type[] parameters)
        {
            IEnumerable<Type> parameterTypes = method.GetParameters().Select(p => p.ParameterType);
            return parameterTypes.SequenceEqual(parameters);
        }
    }
}
