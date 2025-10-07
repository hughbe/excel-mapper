namespace ExcelMapper.Abstractions;

public interface IDictionaryFactory<TValue>
{
    void Begin(int capacity);
    void Add(string key, TValue? value);
    object End();
    void Reset();
}
