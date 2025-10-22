using ExcelMapper.Abstractions;

namespace ExcelMapper.Tests;

public class CantReadTests
{
    [Fact]
    public void CantRead_NoSuchColumnIndexProperty_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnIndexPropertyClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for column at index {int.MaxValue}", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchColumnIndexPropertyClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        public string Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_NoSuchColumnIndexField_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnIndexFieldClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for column at index {int.MaxValue}", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchColumnIndexFieldClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        public string Member = default!;
    }

    [Fact]
    public void CantRead_NoSuchColumnNameProperty_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnNamePropertyClass>());
        Assert.StartsWith("Could not read value for member \"Member\" for column \"NoSuchColumn\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchColumnNamePropertyClass
    {
        [ExcelColumnName("NoSuchColumn")]
        public string Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_NoSuchColumnNameField_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnNameFieldClass>());
        Assert.StartsWith("Could not read value for member \"Member\" for column \"NoSuchColumn\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchColumnNameFieldClass
    {
        [ExcelColumnName("NoSuchColumn")]
        public string Member = default!;
    }

    [Fact]
    public void CantRead_NoSuchColumnMatchingProperty_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnMatchingPropertyClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" (no columns matching)", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    [Fact]
    public void CantRead_NoSuchColumnMatchingPropertyNoHeading_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnMatchingPropertyClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" (no columns matching)", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchColumnMatchingPropertyClass
    {
        [ExcelColumnMatching(typeof(NeverMatcher))]
        public string Member { get; set; } = default!;
    }

    private class NeverMatcher : IExcelColumnMatcher
    {
        public bool ColumnMatches(ExcelSheet sheet, int columnIndex) => false;
    }

    [Fact]
    public void CantRead_NoSuchColumnMatchingField_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnMatchingFieldClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" (no columns matching)", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    [Fact]
    public void CantRead_NoSuchColumnMatchingFieldNoHeading_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnMatchingFieldClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" (no columns matching)", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchColumnMatchingFieldClass
    {
        [ExcelColumnMatching(typeof(NeverMatcher))]
        public string Member = default!;
    }

    [Fact]
    public void CantRead_SplitNoSuchColumnIndexProperty_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchSplitColumnIndexPropertyClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for column at index {int.MaxValue}", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchSplitColumnIndexPropertyClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        public IEnumerable<string> Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_SplitNoSuchColumnIndexField_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchSplitColumnIndexFieldClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for column at index {int.MaxValue}", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchSplitColumnIndexFieldClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        public IEnumerable<string> Member = default!;
    }

    [Fact]
    public void CantRead_SplitNoSuchColumnNameProperty_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchSplitColumnNamePropertyClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for column \"NoSuchColumn\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchSplitColumnNamePropertyClass
    {
        [ExcelColumnName("NoSuchColumn")]
        public IEnumerable<string> Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_SplitNoSuchColumnNameField_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchSplitColumnNameFieldClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for column \"NoSuchColumn\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchSplitColumnNameFieldClass
    {
        [ExcelColumnName("NoSuchColumn")]
        public IEnumerable<string> Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_EnumerableNoSuchColumnIndexProperty_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchEnumerableColumnIndexPropertyClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for columns at indices 1, {int.MaxValue}", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchEnumerableColumnIndexPropertyClass
    {
        [ExcelColumnIndices(1, int.MaxValue)]
        public IEnumerable<string> Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_EnumerableNoSuchColumnIndexField_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchEnumerableColumnIndexFieldClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for columns at indices 1, {int.MaxValue}", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchEnumerableColumnIndexFieldClass
    {
        [ExcelColumnIndices(1, int.MaxValue)]
        public IEnumerable<string> Member = default!;
    }

    [Fact]
    public void CantRead_EnumerableNoSuchColumnNameProperty_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchEnumerableColumnNamePropertyClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for columns \"Value\", \"NoSuchColumn\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchEnumerableColumnNamePropertyClass
    {
        [ExcelColumnNames("Value", "NoSuchColumn")]
        public IEnumerable<string> Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_EnumerableNoSuchColumnNameField_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchEnumerableColumnNameFieldClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for columns \"Value\", \"NoSuchColumn\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchEnumerableColumnNameFieldClass
    {
        [ExcelColumnNames("Value", "NoSuchColumn")]
        public IEnumerable<string> Member = default!;
    }

    [Fact]
    public void CantRead_ExceptionThrownProperty_ThrowsCorrectException()
    {
        var exception = new InvalidOperationException("Test exception");
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<IntPropertyClass>(c =>
        {
            c.Map(m => m.Member)
                .WithColumnName("Int Value")
                .WithMappers(new ExceptionThrowingMapper(exception));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntPropertyClass>());
        Assert.Same(exception, ex.InnerException);
        Assert.StartsWith("Cannot assign \"1\" to member \"Member\" of type \"System.Int32\" in column \"Int Value\" on row 0 in sheet \"Primitives\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(0, ex.ColumnIndex);
    }

    [Fact]
    public void CantRead_ExceptionThrownField_ThrowsCorrectException()
    {
        var exception = new InvalidOperationException("Test exception");
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<IntFieldClass>(c =>
        {
            c.Map(m => m.Member)
                .WithColumnName("Int Value")
                .WithMappers(new ExceptionThrowingMapper(exception));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntFieldClass>());
        Assert.Same(exception, ex.InnerException);
        Assert.StartsWith("Cannot assign \"1\" to member \"Member\" of type \"System.Int32\" in column \"Int Value\" on row 0 in sheet \"Primitives\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(0, ex.ColumnIndex);
    }

    [Fact]
    public void CantRead_NullExceptionProperty_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<IntPropertyClass>(c =>
        {
            c.Map(m => m.Member)
                .WithColumnName("Int Value")
                .WithMappers(new ExceptionThrowingMapper(null!));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntPropertyClass>());
        Assert.Null(ex.InnerException);
        Assert.StartsWith("Cannot assign \"1\" to member \"Member\" of type \"System.Int32\" in column \"Int Value\" on row 0 in sheet \"Primitives\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(0, ex.ColumnIndex);
    }

    [Fact]
    public void CantRead_NullExceptionField_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<IntFieldClass>(c =>
        {
            c.Map(m => m.Member)
                .WithColumnName("Int Value")
                .WithMappers(new ExceptionThrowingMapper(null!));
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntFieldClass>());
        Assert.Null(ex.InnerException);
        Assert.StartsWith("Cannot assign \"1\" to member \"Member\" of type \"System.Int32\" in column \"Int Value\" on row 0 in sheet \"Primitives\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(0, ex.ColumnIndex);
    }

    private class IntPropertyClass
    {
        public int Member { get; set; } = default!;
    }

    private class IntFieldClass
    {
        public int Member { get; set; } = default!;
    }

    private class ExceptionThrowingMapper(Exception exception) : ICellMapper
    {
        private readonly Exception _exception = exception;

        public CellMapperResult Map(ReadCellResult readResult) => CellMapperResult.Invalid(_exception);
    }

    [Fact]
    public void CantRead_NoValidMappersProperty_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<IntPropertyClass>(c =>
        {
            var map = c.Map(m => m.Member)
                .WithColumnName("Int Value")
                .WithMappers(new IgnoreCellMapper());
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntPropertyClass>());
        Assert.Null(ex.InnerException);
        Assert.StartsWith("Cannot assign \"1\" to member \"Member\" of type \"System.Int32\" in column \"Int Value\" on row 0 in sheet \"Primitives\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(0, ex.ColumnIndex);
    }

    [Fact]
    public void CantRead_NoValidMappersField_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<IntFieldClass>(c =>
        {
            var map = c.Map(m => m.Member)
                .WithColumnName("Int Value")
                .WithMappers(new IgnoreCellMapper());
        });

        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntFieldClass>());
        Assert.Null(ex.InnerException);
        Assert.StartsWith("Cannot assign \"1\" to member \"Member\" of type \"System.Int32\" in column \"Int Value\" on row 0 in sheet \"Primitives\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(0, ex.ColumnIndex);
    }

    private class IgnoreCellMapper : ICellMapper
    {
        public CellMapperResult Map(ReadCellResult readResult) => CellMapperResult.Ignore();
    }
}
