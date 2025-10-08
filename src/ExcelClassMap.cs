using System;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

public class ExcelClassMap : IMap
{
    public ExcelClassMap(Type type) : this(type, FallbackStrategy.ThrowIfPrimitive)
    {
    }

    public ExcelClassMap(Type type, FallbackStrategy emptyValueStrategy)
    {
        Type = type ?? throw new ArgumentNullException(nameof(type));

        if (!Enum.IsDefined(emptyValueStrategy))
        {
            throw new ArgumentException($"Invalid value \"{emptyValueStrategy}\".", nameof(emptyValueStrategy));
        }
        EmptyValueStrategy = emptyValueStrategy;
    }

    public Type Type { get; }

    public ExcelPropertyMapCollection Properties { get; } = [];

    /// <summary>
    /// Gets the default strategy to use when the value of a cell is empty.
    /// </summary>
    public FallbackStrategy EmptyValueStrategy { get; private protected set; }

    public virtual bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
    {
        var instance = Activator.CreateInstance(Type)!;
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
