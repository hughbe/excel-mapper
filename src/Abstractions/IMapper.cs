
namespace ExcelMapper.Abstractions;

public interface IMapper : IMap
{
    CellValueMapperCollection Mappers { get; }
}
