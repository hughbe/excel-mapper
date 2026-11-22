using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.ExceptionServices;

namespace ExcelMapper.Utilities;

internal static class ReflectionUtilities
{
    /// <summary>
    /// Whether <paramref name="type"/> implements <paramref name="interfaceType"/>. An interface
    /// doesn't implement itself.
    /// </summary>
    public static bool ImplementsInterface([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.Interfaces)] this Type type, Type interfaceType)
        => Array.IndexOf(type.GetInterfaces(), interfaceType) >= 0;

    /// <summary>
    /// Whether <paramref name="type"/> is or implements a closed <paramref name="genericInterfaceType"/>,
    /// such as <c>IEnumerable&lt;T&gt;</c>, giving its first type argument.
    /// </summary>
    public static bool ImplementsGenericInterface(
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.Interfaces)] this Type type,
        Type genericInterfaceType,
        [NotNullWhen(true)] out Type? elementType)
    {
        bool IsGenericInterface(Type interfaceType, [NotNullWhen(true)] out Type? typeArgument)
        {
            typeArgument = interfaceType.IsGenericType && interfaceType.GetGenericTypeDefinition() == genericInterfaceType
                ? interfaceType.GenericTypeArguments[0]
                : null;
            return typeArgument is not null;
        }

        // The type may be the interface itself.
        if (type.IsInterface && IsGenericInterface(type, out elementType))
        {
            return true;
        }

        foreach (var interfaceType in type.GetInterfaces())
        {
            if (IsGenericInterface(interfaceType, out elementType))
            {
                return true;
            }
        }

        elementType = null;
        return false;
    }

    /// <summary>
    /// The type of a mapped member, which is always a property or field.
    /// </summary>
    public static Type MemberType(this MemberInfo member)
        => member is PropertyInfo property ? property.PropertyType : ((FieldInfo)member).FieldType;

    public static Type GetNullableTypeOrThis(this Type type, out bool isNullable)
    {
        isNullable = type.IsNullable();
        return isNullable ? type.GenericTypeArguments[0] : type;
    }

    public static bool IsNullable(this Type type)
    {
        return Nullable.GetUnderlyingType(type) is not null;
    }

    /// <summary>
    /// Gets the element type or the IEnumerable&lt;T&gt; type of the given type.
    /// </summary>
    /// <param name="type">The type to get the element type of.</param>
    /// <returns>The element type or IEnumerable&lt;T&gt; of the given type, or null if it has none.</returns>
    // Keep in sync with the source generator's TypeWalker.TryGetElementType. See "Auto-mapping and the source generator" in CLAUDE.md.
    public static Type? GetElementTypeOrEnumerableType([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.Interfaces | DynamicallyAccessedMemberTypes.PublicMethods)] this Type type)
    {
        // Array type.
        if (type.IsArray)
        {
            return type.GetElementType()!;
        }

        // Generic lists use type object.
        if (type.ImplementsGenericInterface(typeof(IEnumerable<>), out var elementType))
        {
            return elementType;
        }

        // Non-generic interfaces use type object.
        if (type == typeof(IEnumerable) || type.ImplementsInterface(typeof(IEnumerable)))
        {
            // Some non-generic collections only accept one type, such as StringCollection with
            // Add(string). Use that type so elements are converted to it.
            if (!type.IsInterface && TryGetSingleAddParameterType(type, out var addType))
            {
                return addType;
            }

            return typeof(object);
        }

        return null;
    }
    
    private static bool TryGetSingleAddParameterType([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicMethods)] Type type, [NotNullWhen(true)] out Type? parameterType)
    {
        parameterType = null;
        foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance))
        {
            if (method.Name != "Add" || method.GetParameters() is not { Length: 1 } parameters)
            {
                continue;
            }

            if (parameterType is not null)
            {
                // More than one Add method, so the element type is ambiguous.
                parameterType = null;
                return false;
            }

            parameterType = parameters[0].ParameterType;
        }

        return parameterType is not null && parameterType != typeof(object);
    }

    /// <summary>
    /// Invokes a method through reflection, where code can't be compiled, rethrowing what it throws
    /// rather than a <see cref="TargetInvocationException"/>.
    /// </summary>
    public static object InvokeUnwrapped(this MethodInfo method, object? obj, params object?[] parameters)
    {
        try
        {
            return method.Invoke(obj, parameters)!;
        }
        catch (TargetInvocationException ex)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException!).Throw();
            throw; // Will never be hit, but compiler doesn't know that.
        }
    }
}
