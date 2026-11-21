using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Security;
using System.Text;

namespace ExcelMapper.Tests;

/// <summary>
/// Builds small .xlsx workbooks in memory so tests can use typed Excel cells (numbers, dates,
/// booleans, formulas, errors and formatted values) without binary fixture files.
/// </summary>
public static class XlsxBuilder
{
    /// <summary>A number displayed with an Excel number format, for example "0.00%" or "00000".</summary>
    public record Formatted(double Value, string Format);

    /// <summary>A formula with its cached result.</summary>
    public record Formula(string Expression, object Result);

    /// <summary>An error value such as "#N/A" or "#DIV/0!".</summary>
    public record Error(string Value);

    /// <summary>A duration displayed with the [h]:mm:ss format.</summary>
    public record Duration(TimeSpan Value);

    /// <summary>
    /// Creates an importer for a single-sheet workbook. Each row is an array of cell values: strings,
    /// numbers, booleans, DateTime (date and time), DateOnly (date), TimeOnly (time), null (empty),
    /// or one of the records above.
    /// </summary>
    public static ExcelImporter Create(params object?[][] rows) => new(CreateStream(rows));

    public static MemoryStream CreateStream(params object?[][] rows)
    {
        // Built-in formats: 14 = date, 22 = date and time, 21 = time, 46 = [h]:mm:ss duration.
        // Custom formats are numbered from 164.
        var customFormats = new List<string>();
        var styles = new List<int> { 0, 14, 22, 21, 46 };
        int StyleFor(string format)
        {
            var index = customFormats.IndexOf(format);
            if (index == -1)
            {
                customFormats.Add(format);
                index = customFormats.Count - 1;
            }

            var styleIndex = styles.IndexOf(164 + index);
            if (styleIndex == -1)
            {
                styles.Add(164 + index);
                styleIndex = styles.Count - 1;
            }

            return styleIndex;
        }

        var sheet = new StringBuilder();
        sheet.Append("<sheetData>");
        for (int r = 0; r < rows.Length; r++)
        {
            sheet.Append($"<row r=\"{r + 1}\">");
            for (int c = 0; c < rows[r].Length; c++)
            {
                var reference = ExcelRangeHelpers.ColumnLetters(c) + (r + 1);
                sheet.Append(CellXml(reference, rows[r][c], StyleFor));
            }

            sheet.Append("</row>");
        }

        sheet.Append("</sheetData>");

        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            Add(archive, "[Content_Types].xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                </Types>
                """);
            Add(archive, "_rels/.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);
            Add(archive, "xl/workbook.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """);
            Add(archive, "xl/_rels/workbook.xml.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """);

            var numFmts = new StringBuilder();
            for (int i = 0; i < customFormats.Count; i++)
            {
                numFmts.Append($"<numFmt numFmtId=\"{164 + i}\" formatCode=\"{SecurityElement.Escape(customFormats[i])}\"/>");
            }

            var xfs = new StringBuilder();
            foreach (var numFmtId in styles)
            {
                xfs.Append($"<xf numFmtId=\"{numFmtId}\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"{(numFmtId == 0 ? "" : " applyNumberFormat=\"1\"")}/>");
            }

            Add(archive, "xl/styles.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <numFmts count="{customFormats.Count}">{numFmts}</numFmts>
                <fonts count="1"><font/></fonts>
                <fills count="1"><fill/></fills>
                <borders count="1"><border/></borders>
                <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
                <cellXfs count="{styles.Count}">{xfs}</cellXfs>
                </styleSheet>
                """);
            Add(archive, "xl/worksheets/sheet1.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{sheet}</worksheet>
                """);
        }

        stream.Position = 0;
        return stream;
    }

    private static string CellXml(string reference, object? value, Func<string, int> styleFor)
    {
        static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);

        return value switch
        {
            null => "",
            string s => $"<c r=\"{reference}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{SecurityElement.Escape(s)}</t></is></c>",
            bool b => $"<c r=\"{reference}\" t=\"b\"><v>{(b ? 1 : 0)}</v></c>",
            DateTime d when d.TimeOfDay == TimeSpan.Zero => $"<c r=\"{reference}\" s=\"1\"><v>{Number(d.ToOADate())}</v></c>",
            DateTime d => $"<c r=\"{reference}\" s=\"2\"><v>{Number(d.ToOADate())}</v></c>",
            DateOnly d => $"<c r=\"{reference}\" s=\"1\"><v>{Number(d.ToDateTime(TimeOnly.MinValue).ToOADate())}</v></c>",
            TimeOnly t => $"<c r=\"{reference}\" s=\"3\"><v>{Number(t.ToTimeSpan().TotalDays)}</v></c>",
            Duration d => $"<c r=\"{reference}\" s=\"4\"><v>{Number(d.Value.TotalDays)}</v></c>",
            Formatted f => $"<c r=\"{reference}\" s=\"{styleFor(f.Format)}\"><v>{Number(f.Value)}</v></c>",
            Error e => $"<c r=\"{reference}\" t=\"e\"><v>{SecurityElement.Escape(e.Value)}</v></c>",
            Formula { Result: string s } f => $"<c r=\"{reference}\" t=\"str\"><f>{SecurityElement.Escape(f.Expression)}</f><v>{SecurityElement.Escape(s)}</v></c>",
            Formula f => $"<c r=\"{reference}\"><f>{SecurityElement.Escape(f.Expression)}</f><v>{Number(Convert.ToDouble(f.Result, CultureInfo.InvariantCulture))}</v></c>",
            IConvertible n => $"<c r=\"{reference}\"><v>{Number(Convert.ToDouble(n, CultureInfo.InvariantCulture))}</v></c>",
            _ => throw new ArgumentException($"Unsupported cell value {value} of type {value.GetType()}.", nameof(value))
        };
    }

    private static void Add(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content.Trim());
    }

    private static class ExcelRangeHelpers
    {
        public static string ColumnLetters(int index)
        {
            var letters = "";
            for (var n = index + 1; n > 0; n = (n - 1) / 26)
            {
                letters = (char)('A' + (n - 1) % 26) + letters;
            }

            return letters;
        }
    }
}
