using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a multiples cells given a list of column names or a predicate matching the column name.
/// </summary>
public sealed class ColumnsMatchingReaderFactory : ICellReaderFactory, ICellsReaderFactory, IColumnIndicesProviderCellReaderFactory, IColumnNamesProviderCellReaderFactory
{
    public IExcelColumnMatcher Matcher { get; }

    /// <summary>
    /// Constructs a reader that reads the value of multiple cells given the predicate matching the column name.
    /// </summary>
    /// <param name="predicate">The predicate containing the column name to read.</param>
    public ColumnsMatchingReaderFactory(IExcelColumnMatcher matcher)
    {
        Matcher = matcher ?? throw new ArgumentNullException(nameof(matcher));
    }

    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

        for (var columnIndex = 0; columnIndex < sheet.NumberOfColumns; columnIndex++)
        {
            if (Matcher.ColumnMatches(sheet, columnIndex))
            {
                return new ColumnIndexReader(columnIndex);
            }
        }

        return null;
    }

    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

        var indices = new List<int>();
        for (var columnIndex = 0; columnIndex < sheet.NumberOfColumns; columnIndex++)
        {
            if (Matcher.ColumnMatches(sheet, columnIndex))
            {
                indices.Add(columnIndex);
            }
        }

        if (indices.Count == 0)
        {
            return null;
        }

        return new ColumnIndicesReader(indices);
    }

    public string[] GetColumnNames(ExcelSheet sheet)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        if (sheet.Heading == null)
        {
            return null!;
        }

        var names = new List<string>();
        for (var columnIndex = 0; columnIndex < sheet.NumberOfColumns; columnIndex++)
        {
            if (Matcher.ColumnMatches(sheet, columnIndex))
            {
                names.Add(sheet.Heading.GetColumnName(columnIndex)!);
            }
        }

        return names.ToArray();
    }

    public int[] GetColumnIndices(ExcelSheet sheet)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

        var indices = new List<int>();
        for (var columnIndex = 0; columnIndex < sheet.NumberOfColumns; columnIndex++)
        {
            if (Matcher.ColumnMatches(sheet, columnIndex))
            {
                indices.Add(columnIndex);
            }
        }

        return indices.ToArray();
    }
}