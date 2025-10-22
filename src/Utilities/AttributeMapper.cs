using System.Linq;
using System.Reflection;
using ExcelMapper.Readers;

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
}
