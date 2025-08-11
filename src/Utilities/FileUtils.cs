using System;

namespace ExcelMapper.Utilities;

/// <summary>
/// Contains file extension constants for supported file formats.
/// </summary>
public static class FileUtils
{
    #region CSV Extensions
    /// <summary>
    /// Comma-separated values file extension.
    /// </summary>
    public const string Csv = ".csv";
    #endregion

    #region Excel Extensions
    /// <summary>
    /// Excel 97-2003 workbook file extension.
    /// </summary>
    public const string Xls = ".xls";
    
    /// <summary>
    /// Excel workbook file extension.
    /// </summary>
    public const string Xlsx = ".xlsx";
    
    /// <summary>
    /// Excel macro-enabled workbook file extension.
    /// </summary>
    public const string Xlsm = ".xlsm";
    
    /// <summary>
    /// Excel binary workbook file extension.
    /// </summary>
    public const string Xlsb = ".xlsb";
    #endregion


    #region Collections
    /// <summary>
    /// All supported CSV file extensions.
    /// </summary>
    public static readonly string[] CsvExtensions = 
    {
        Csv
    };

    /// <summary>
    /// All supported Excel file extensions.
    /// </summary>
    public static readonly string[] ExcelExtensions = 
    {
        Xls,
        Xlsx,
        Xlsm,
        Xlsb
    };

    /// <summary>
    /// All supported file extensions for spreadsheet processing.
    /// </summary>
    public static readonly string[] AllSupportedExtensions =
    [
        Csv,
        Xls,
        Xlsx,
        Xlsm,
        Xlsb
    ];
    #endregion

    #region Helper Methods
    /// <summary>
    /// Determines if the specified extension is a CSV file.
    /// </summary>
    /// <param name="extension">The file extension to check (case-insensitive).</param>
    /// <returns>True if the extension represents a CSV file; otherwise, false.</returns>
    public static bool IsCsvExtension(string extension)
    {
        if (string.IsNullOrWhiteSpace(extension))
            return false;

        return string.Equals(extension, Csv, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Determines if the specified extension is an Excel file.
    /// </summary>
    /// <param name="extension">The file extension to check (case-insensitive).</param>
    /// <returns>True if the extension represents an Excel file; otherwise, false.</returns>
    public static bool IsExcelExtension(string extension)
    {
        if (string.IsNullOrWhiteSpace(extension))
            return false;

        return Array.Exists(ExcelExtensions, 
            ext => string.Equals(ext, extension, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Determines if the specified extension is supported for processing.
    /// </summary>
    /// <param name="extension">The file extension to check (case-insensitive).</param>
    /// <returns>True if the extension is supported; otherwise, false.</returns>
    public static bool IsSupportedExtension(string extension)
    {
        if (string.IsNullOrWhiteSpace(extension))
            return false;

        return Array.Exists(AllSupportedExtensions,
            ext => string.Equals(ext, extension, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Normalizes a file extension by ensuring it starts with a dot and is lowercase.
    /// </summary>
    /// <param name="extension">The file extension to normalize.</param>
    /// <returns>The normalized extension, or null if input is invalid.</returns>
    public static string? NormalizeExtension(string? extension)
    {
        if (string.IsNullOrWhiteSpace(extension))
            return null;

        extension = extension.Trim();
        
        if (!extension.StartsWith("."))
            extension = "." + extension;

        return extension.ToLowerInvariant();
    }
    #endregion
}
