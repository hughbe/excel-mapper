using ExcelDataReader;

namespace ExcelMapper.Reading;

/// <summary>
/// A worksheet of an importer's file: what's fixed about it, shared by its <see cref="ExcelSheet"/>
/// and the views of it that regions are read with (see <see cref="ExcelRegion"/>).
/// </summary>
internal sealed class Worksheet
{
    public Worksheet(IExcelDataReader reader, int index, ExcelImporter importer)
    {
        Reader = reader;
        Name = reader.Name;

        // CSV readers have no concept of sheet visibility and return a null state.
        Visibility = reader.VisibleState switch
        {
            null or "visible" => ExcelSheetVisibility.Visible,
            "hidden" => ExcelSheetVisibility.Hidden,
            _ => ExcelSheetVisibility.VeryHidden,
        };
        Index = index;
        NumberOfColumns = reader.FieldCount;
        Importer = importer;
    }

    public IExcelDataReader Reader { get; }

    public string Name { get; }

    public ExcelSheetVisibility Visibility { get; }

    public int Index { get; }

    public int NumberOfColumns { get; }

    public ExcelImporter Importer { get; }

    public ExcelImporterConfiguration Configuration => Importer.Configuration;

    public RowSource RowSource => Importer.RowSource;

    public void ThrowIfClosed()
    {
        if (Reader.IsClosed)
        {
            throw new ExcelMappingException("The underlying reader is closed.");
        }
    }
}
