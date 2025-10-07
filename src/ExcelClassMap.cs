using System;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

public class ExcelClassMap : IMap
{
    public ExcelClassMap(Type type)
    {
        Type = type ?? throw new ArgumentNullException(nameof(type));
    }

    public Type Type { get; }

    public ExcelPropertyMapCollection Properties { get; } = [];

    public virtual bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
    {
        var instance = Activator.CreateInstance(Type);
        foreach (var property in Properties)
        {
            if (property.Map.TryGetValue(sheet, rowIndex, reader, property.Member, out var value))
            {
                property.SetValueFactory(instance, value);
            }
        }

        result = instance;
        return true;
    }
}
