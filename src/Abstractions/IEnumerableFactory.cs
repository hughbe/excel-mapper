namespace ExcelMapper.Abstractions;

public interface IEnumerableFactory<T>
{
    void Begin(int capacity);
    void Add(T? item);
    object End();
    void Reset();
}
