using System.Data;
using ExcelDataReader;

namespace ExcelMapper.Reading;

/// <summary>
/// Reads a blank row that the importer's reader has already read past: every cell is empty. To
/// tell whether a blank row is at the end of a sheet, <see cref="RowSource"/> reads ahead to the
/// next row with a value, and the blank rows before it are read with this. Everything but the
/// cells comes from the importer's reader, which is on the same sheet.
/// </summary>
internal sealed class BlankRowReader(IExcelDataReader reader) : IExcelDataReader
{
    // The cells: all empty.
    // IDataRecord declares a non-null return, but ExcelDataReader returns null for an empty cell.
    public object GetValue(int i) => null!;

    public bool IsDBNull(int i) => true;

    public string? GetNumberFormatString(int i) => null;

    public string? GetNumberFormatString(int i, IFormatProvider formatProvider) => null;

    public int GetNumberFormatIndex(int i) => -1;

    // The style of the row the importer's reader is on: ExcelMapper doesn't read styles.
    public CellStyle GetCellStyle(int i) => reader.GetCellStyle(i);

    public CellError? GetCellError(int i) => null;

    public int GetValues(object[] values)
    {
        var count = Math.Min(values.Length, FieldCount);
        for (int i = 0; i < count; i++)
        {
            values[i] = DBNull.Value;
        }

        return count;
    }

    // Typed values: a blank cell has none, as IDataRecord's getters throw for a null cell.
    public bool GetBoolean(int i) => throw Blank();

    public byte GetByte(int i) => throw Blank();

    public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) => throw Blank();

    public char GetChar(int i) => throw Blank();

    public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) => throw Blank();

    public IDataReader GetData(int i) => throw Blank();

    public DateTime GetDateTime(int i) => throw Blank();

    public decimal GetDecimal(int i) => throw Blank();

    public double GetDouble(int i) => throw Blank();

    public float GetFloat(int i) => throw Blank();

    public Guid GetGuid(int i) => throw Blank();

    public short GetInt16(int i) => throw Blank();

    public int GetInt32(int i) => throw Blank();

    public long GetInt64(int i) => throw Blank();

    public string GetString(int i) => throw Blank();

    public object this[int i] => DBNull.Value;

    public object this[string name] => throw new NotSupportedException();

    private static InvalidCastException Blank() => new("The cell is empty.");

    // The row and sheet: the importer's reader's.
    public int FieldCount => reader.FieldCount;

    public double RowHeight => reader.RowHeight;

    public string GetDataTypeName(int i) => reader.GetDataTypeName(i);

#if NET8_0_OR_GREATER
    [return: System.Diagnostics.CodeAnalysis.DynamicallyAccessedMembers(System.Diagnostics.CodeAnalysis.DynamicallyAccessedMemberTypes.PublicFields | System.Diagnostics.CodeAnalysis.DynamicallyAccessedMemberTypes.PublicProperties)]
#endif
    public Type GetFieldType(int i) => reader.GetFieldType(i);

    public string GetName(int i) => reader.GetName(i);

    public int GetOrdinal(string name) => reader.GetOrdinal(name);

    public double GetColumnWidth(int i) => reader.GetColumnWidth(i);

    public string Name => reader.Name;

    public string CodeName => reader.CodeName;

    public string VisibleState => reader.VisibleState;

    public HeaderFooter HeaderFooter => reader.HeaderFooter;

    public CellRange[] MergeCells => reader.MergeCells;

    public int ResultsCount => reader.ResultsCount;

    public int RowCount => reader.RowCount;

    public int ActiveSheet => reader.ActiveSheet;

    public bool IsActiveSheet => reader.IsActiveSheet;

    public CellRange Dimension => reader.Dimension;

    public int Depth => reader.Depth;

    public bool IsClosed => reader.IsClosed;

    public int RecordsAffected => reader.RecordsAffected;

    public DataTable? GetSchemaTable() => reader.GetSchemaTable();

    // Moving: only RowSource moves the importer's reader.
    public bool Read() => throw new NotSupportedException();

    public bool NextResult() => throw new NotSupportedException();

    public void Reset() => throw new NotSupportedException();

    public void Close() => throw new NotSupportedException();

    public void Dispose()
    {
    }
}
