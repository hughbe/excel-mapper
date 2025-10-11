namespace ExcelMapper.Abstractions;

public interface IDictionaryFactory<TKey, TValue>
{
    void Begin(int count);
    void Add(TKey key, TValue? value);
    object End();
    void Reset();
}
