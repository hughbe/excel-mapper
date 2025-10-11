namespace ExcelMapper.Abstractions;

public interface IMultidimensionalArrayFactory<T>
{
    void Begin(int[] lengths);
    void Set(int[] indices, T? item);
    object End();
    void Reset();
}
