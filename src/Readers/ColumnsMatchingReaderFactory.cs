using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a multiples cells given a list of column names or a predicate matching the column name.
/// </summary>
public sealed class ColumnsMatchingReaderFactory : ICellReaderFactory, ICellsReaderFactory
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
}