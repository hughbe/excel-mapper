
namespace ExcelMapper.Readers;

/// <summary>
/// A factory that combines multiple cell reader factories.
/// </summary>
public class CompositeCellsReaderFactory : ICellReaderFactory, ICellsReaderFactory, IColumnNameProviderCellReaderFactory, IColumnIndexProviderCellReaderFactory, IColumnNamesProviderCellReaderFactory, IColumnIndicesProviderCellReaderFactory
{
    /// <summary>
    /// Gets the cell reader factories.
    /// </summary>
    public IReadOnlyList<ICellReaderFactory> Factories { get; }

    /// <summary>
    /// Initializes a new instance of <see cref="CompositeCellsReaderFactory"/>.
    /// </summary>
    /// <param name="factories">The cell reader factories.</param>
    /// <exception cref="ArgumentException">Thrown when the factories array is empty or contains null values.</exception>
    public CompositeCellsReaderFactory(params IReadOnlyList<ICellReaderFactory> factories)
    {
        ArgumentNullException.ThrowIfNull(factories, nameof(factories));
        if (factories.Count == 0)
        {
            throw new ArgumentException("At least one factory must be provided.", nameof(factories));
        }
        foreach (var factory in factories)
        {
            if (factory == null)
            {
                throw new ArgumentException("Factories cannot contain null values.", nameof(factories));
            }
        }

        Factories = factories;
    }

    /// <inheritdoc/>
    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        foreach (var factory in Factories)
        {
            var reader = factory.GetCellReader(sheet);
            if (reader != null)
            {
                return reader;
            }
        }

        return null;
    }

    /// <inheritdoc/>
    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        var readers = new List<ICellReader>(Factories.Count);
        for (int i = 0; i < Factories.Count; i++)
        {
            var reader = Factories[i].GetCellReader(sheet);
            if (reader != null)
            {
                readers.Add(reader);
            }
        }

        if (readers.Count == 0)
        {
            return null;
        }

        return new CompositeCellsReader(readers);
    }

    public string GetColumnName(ExcelSheet sheet)
    {
        foreach (var factory in Factories)
        {
            if (factory is IColumnNameProviderCellReaderFactory nameProviderFactory)
            {
                var columnName = nameProviderFactory.GetColumnName(sheet);
                if (!string.IsNullOrEmpty(columnName))
                {
                    return columnName;
                }
            }
        }

        return string.Empty;
    }

    public int? GetColumnIndex(ExcelSheet sheet)
    {
        foreach (var factory in Factories)
        {
            if (factory is IColumnIndexProviderCellReaderFactory indexProviderFactory)
            {
                var columnIndex = indexProviderFactory.GetColumnIndex(sheet);
                if (columnIndex != null && columnIndex != -1)
                {
                    return columnIndex;
                }
            }
        }

        return null;
    }

    public IReadOnlyList<string>? GetColumnNames(ExcelSheet sheet)
    {
        var names = new List<string>(Factories.Count);
        foreach (var factory in Factories)
        {
            if (factory is IColumnNameProviderCellReaderFactory namesProviderFactory)
            {
                var columnName = namesProviderFactory.GetColumnName(sheet);
                if (!string.IsNullOrEmpty(columnName))
                {
                    names.Add(columnName);
                }
            }
        }

        return names.Count > 0 ? names : null;
    }

    public IReadOnlyList<int>? GetColumnIndices(ExcelSheet sheet)
    {
        var indices = new List<int>(Factories.Count);
        foreach (var factory in Factories)
        {
            if (factory is IColumnIndexProviderCellReaderFactory indicesProviderFactory)
            {
                var columnIndex = indicesProviderFactory.GetColumnIndex(sheet);
                if (columnIndex != null && columnIndex != -1)
                {
                    indices.Add(columnIndex.Value);
                }
            }
        }

        return indices.Count > 0 ? indices : null;
    }
}
