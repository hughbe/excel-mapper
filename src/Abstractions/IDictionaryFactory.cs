namespace ExcelMapper.Abstractions;

public interface IDictionaryFactory<TValue>
{
    void Begin(int count);
    void Add(string key, TValue? value);
    object End();
    void Reset();
}
