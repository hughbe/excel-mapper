using System.Linq;
using System.Reflection;
using ExcelMapper.Fallbacks;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;

namespace ExcelMapper.Utilities;

internal static class MemberMapper
{
    internal static ICellsReaderFactory? GetDefaultCellsReaderFactory(MemberInfo? member)
    {
        // If no member was specified, read all the cells.
        if (member == null)
        {
            return new AllColumnNamesReaderFactory();
        }

        // [ExcelColumnNames] attributes represent multiple columns.
        var columnNamesAttribute = member.GetCustomAttribute<ExcelColumnNamesAttribute>();
        if (columnNamesAttribute != null)
        {
            return new ColumnNamesReaderFactory(columnNamesAttribute.Names);
        }

        // [ExcelColumnsMatchingAttribute] attributes represent multiple columns.
        var columnNameMatchingAttribute = member.GetCustomAttribute<ExcelColumnsMatchingAttribute>();
        if (columnNameMatchingAttribute != null)
        {
            var matcher = (IExcelColumnMatcher)Activator.CreateInstance(columnNameMatchingAttribute.Type, columnNameMatchingAttribute.ConstructorArguments)!;
            return new ColumnsMatchingReaderFactory(matcher);
        }

        // [ExcelColumnIndices] attributes represents multiple columns.
        var columnIndicesAttribute = member.GetCustomAttribute<ExcelColumnIndicesAttribute>();
        if (columnIndicesAttribute != null)
        {
            return new ColumnIndicesReaderFactory(columnIndicesAttribute.Indices);
        }

        return null;
    }

    internal static ICellReaderFactory GetDefaultCellReaderFactory(MemberInfo member)
    {
        var columnNameAttributes = member.GetCustomAttributes<ExcelColumnNameAttribute>().ToArray();
        // A single [ExcelColumnName] attribute represents one column.
        if (columnNameAttributes.Length == 1)
        {
            return new ColumnNameReaderFactory(columnNameAttributes[0].Name);
        }
        // Multiple [ExcelColumnName] attributes still represents one column, but multiple options.
        else if (columnNameAttributes.Length > 1)
        {
            return new ColumnNamesReaderFactory([.. columnNameAttributes.Select(c => c.Name)]);
        }

        // [ExcelColumnNames] attributes still represents one column, but multiple options.
        var columnNamesAttribute = member.GetCustomAttribute<ExcelColumnNamesAttribute>();
        if (columnNamesAttribute != null)
        {
            return new ColumnNamesReaderFactory(columnNamesAttribute.Names);
        }

        // A single [ExcelColumnNameMatching] attributes still represents one column, but multiple options.
        var columnNameMatchingAttribute = member.GetCustomAttribute<ExcelColumnMatchingAttribute>();
        if (columnNameMatchingAttribute != null)
        {
            var matcher = (IExcelColumnMatcher)Activator.CreateInstance(columnNameMatchingAttribute.Type, columnNameMatchingAttribute.ConstructorArguments)!;
            return new ColumnsMatchingReaderFactory(matcher);
        }

        // A single [ExcelColumnIndex] attribute represents one column.
        var colummnIndexAttributes = member.GetCustomAttributes<ExcelColumnIndexAttribute>().ToArray();
        if (colummnIndexAttributes.Length == 1)
        {
            return new ColumnIndexReaderFactory(colummnIndexAttributes[0].Index);
        }
        // Multiple [ExcelColumnIndex] attributes still represents one column, but multiple options.
        else if (colummnIndexAttributes.Length > 1)
        {
            return new ColumnIndicesReaderFactory([.. colummnIndexAttributes.Select(c => c.Index)]);
        }

        // [ExcelColumnIndices] attributes still represents one column, but multiple options.
        var columnIndicesAttribute = member.GetCustomAttribute<ExcelColumnIndicesAttribute>();
        if (columnIndicesAttribute != null)
        {
            return new ColumnIndicesReaderFactory(columnIndicesAttribute.Indices);
        }

        return new ColumnNameReaderFactory(member.Name);
    }

    internal static IFallbackItem? GetDefaultEmptyValueFallback(MemberInfo member)
    {
        // If the member has a ExcelDefaultValue attribute, add the fallback.
        if (member.GetCustomAttribute<ExcelDefaultValueAttribute>() is { } defaultValueAttribute)
        {
            return new FixedValueFallback(defaultValueAttribute.Value);
        }

        // If the member has a ExcelEmptyFallback attribute, add the fallback.
        if (member.GetCustomAttribute<ExcelEmptyFallbackAttribute>() is { } emptyFallback)
        {
            return (IFallbackItem)Activator.CreateInstance(emptyFallback.Type, emptyFallback.ConstructorArguments)!;
        }

        return null;
    }

    internal static IFallbackItem? GetDefaultInvalidValueFallback(MemberInfo member)
    {
        // If the member has a ExcelInvalidValue attribute, add the fallback.
        if (member.GetCustomAttribute<ExcelInvalidValueAttribute>() is { } invalidValueAttribute)
        {
            return new FixedValueFallback(invalidValueAttribute.Value);
        }

        // If the member has a ExcelInvalidFallback attribute, add the fallback.
        if (member.GetCustomAttribute<ExcelInvalidFallbackAttribute>() is { } invalidFallback)
        {
            return (IFallbackItem)Activator.CreateInstance(invalidFallback.Type, invalidFallback.ConstructorArguments)!;
        }

        return null;
    }

    internal static void AddTransformers(IValuePipeline pipeline, MemberInfo member)
    {
        // If the member has a ExcelTrimString attribute, add the transformer.
        if (member.GetCustomAttribute<ExcelTrimStringAttribute>() is { } trimStringAttribute)
        {
            pipeline.Transformers.Add(new TrimStringCellTransformer());
        }

        // If the member has ExcelTransformer attributes, add the transformers.
        if (member.GetCustomAttributes<ExcelTransformerAttribute>() is { } transformAttributes)
        {
            foreach (var transformAttribute in transformAttributes)
            {
                var transformer = (ICellTransformer)Activator.CreateInstance(transformAttribute.Type, transformAttribute.ConstructorArguments)!;
                pipeline.Transformers.Add(transformer);
            }
        }
    }

    internal static void ApplyMemberAttributes(IToOneMap map, MemberInfo member)
    {
        // If a member has an ExcelOptional attribute, mark it as optional.
        if (Attribute.IsDefined(member, typeof(ExcelOptionalAttribute)))
        {
            map.Optional = true;
        }
        // If a member has an ExcelPreserveFormatting attribute, mark it as preserving formatting.
        if (Attribute.IsDefined(member, typeof(ExcelPreserveFormattingAttribute)))
        {
            map.PreserveFormatting = true;
        }
    }
}
