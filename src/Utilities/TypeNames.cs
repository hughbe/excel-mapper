using System.Collections.Concurrent;
using System.Text;

namespace ExcelMapper.Utilities;

/// <summary>
/// Names types in error messages as they're written in C#: <c>int</c>, <c>int?</c>,
/// <c>List&lt;string&gt;</c> and <c>Employee</c>, rather than <c>System.Int32</c>,
/// <c>System.Nullable`1[System.Int32]</c> and <c>MyApp.Models.Employee</c>.
/// </summary>
internal static class TypeNames
{
    // Names are only needed for errors, but an invalid cell can be reported for every row, so
    // each type is named once.
    private static readonly ConcurrentDictionary<Type, string> s_names = new();

    public static string Get(Type type) => s_names.TryGetValue(type, out var name) ? name : s_names.GetOrAdd(type, Create(type));

    private static string Create(Type type)
    {
        if (GetKeyword(type) is { } keyword)
        {
            return keyword;
        }

        if (Nullable.GetUnderlyingType(type) is { } underlyingType)
        {
            return Get(underlyingType) + "?";
        }

        if (type.IsArray)
        {
            var rank = type.GetArrayRank();
            return Get(type.GetElementType()!) + "[" + new string(',', rank - 1) + "]";
        }

        // A nested type is named after the type it's in, as in C#.
        var name = type.IsNested && !type.IsGenericParameter ? Get(type.DeclaringType!) + "." + type.Name : type.Name;
        if (!type.IsGenericType)
        {
            return name;
        }

        var tick = name.IndexOf('`');
        var builder = new StringBuilder(tick < 0 ? name : name.Substring(0, tick)).Append('<');
        var arguments = type.GetGenericArguments();
        for (int i = 0; i < arguments.Length; i++)
        {
            builder.Append(i == 0 ? string.Empty : ", ").Append(Get(arguments[i]));
        }

        return builder.Append('>').ToString();
    }

    private static string? GetKeyword(Type type)
        => Type.GetTypeCode(type) switch
        {
            _ when type.IsEnum => null,
            TypeCode.Boolean => "bool",
            TypeCode.Byte => "byte",
            TypeCode.SByte => "sbyte",
            TypeCode.Char => "char",
            TypeCode.Int16 => "short",
            TypeCode.UInt16 => "ushort",
            TypeCode.Int32 => "int",
            TypeCode.UInt32 => "uint",
            TypeCode.Int64 => "long",
            TypeCode.UInt64 => "ulong",
            TypeCode.Single => "float",
            TypeCode.Double => "double",
            TypeCode.Decimal => "decimal",
            TypeCode.String => "string",
            _ when type == typeof(object) => "object",
            _ => null,
        };
}
