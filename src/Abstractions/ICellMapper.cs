namespace ExcelMapper.Abstractions;

public interface ICellMapper
{
    CellMapperResult Map(ReadCellResult readResult);
}
