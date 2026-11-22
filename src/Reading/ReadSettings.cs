using System.Globalization;

namespace ExcelMapper.Reading;

/// <summary>
/// The importer's settings as a sheet reads them: taken from <see cref="ExcelImporterConfiguration"/>
/// when the sheet starts reading (its heading or first row), so that changing the configuration
/// part way through doesn't change how the rest of the sheet is read. Sheets read afterwards get
/// the new settings.
/// </summary>
internal sealed class ReadSettings(ExcelImporterConfiguration configuration)
{
    public CultureInfo Culture { get; } = configuration.Culture;

    public bool SkipBlankLines { get; } = configuration.SkipBlankLines;

    public bool ValidateDataAnnotations { get; } = configuration.ValidateDataAnnotations;

    public bool IgnoreSeparatorsInColumnNames { get; } = configuration.IgnoreSeparatorsInColumnNames;

    public int MaxColumnsPerSheet { get; } = configuration.MaxColumnsPerSheet;
}
