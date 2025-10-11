namespace ExcelMapper.Abstractions;

public interface IEnumerableFactory<T>
{
    void Begin(int count);
    void Add(T? item);
    void Set(int index, T? item);
    object End();
    void Reset();
}
