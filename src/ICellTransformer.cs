namespace ExcelMapper;

/// <summary>
/// Changes a cell's value before it's converted, for example to trim or clean up text. Set it with
/// <c>WithTransformers(...)</c> or <see cref="ExcelTransformerAttribute"/>. Transformers run in
/// order, before the empty check: a transformer that returns an empty value makes the cell empty.
/// </summary>
/// <remarks>
/// A transformer can be used for many cells at once on different threads, so it mustn't change its
/// own state.
/// </remarks>
public interface ICellTransformer
{
    /// <summary>
    /// Transforms a cell's value.
    /// </summary>
    /// <param name="value">The cell's value, or the previous transformer's result.</param>
    /// <param name="context">Where the cell is, and the member it's mapped to.</param>
    /// <returns>The value to convert: <paramref name="value"/> to leave it unchanged, or a new
    /// <see cref="CellValue"/>, such as <c>new CellValue(text)</c>.</returns>
    CellValue Transform(CellValue value, CellContext context);
}
