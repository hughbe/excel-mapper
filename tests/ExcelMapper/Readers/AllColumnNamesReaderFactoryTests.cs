using ExcelMapper.Abstractions;
using ExcelMapper.Tests;

namespace ExcelMapper.Readers.Tests;

public class AllColumnNamesValueReaderTests
{
    [Fact]
    public void GetReader_InvokeCanRead_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new AllColumnNamesReaderFactory();
        var reader = factory.GetCellsReader(sheet);
        Assert.NotNull(reader);
        for (var i = 0; i < 2; i++)
        {
            Assert.True(reader.Start(importer.Reader, false, out var count));
            Assert.Equal(1, count);
            var values = new List<string?>();
            while (reader.TryGetNext(out var result))
            {
                values.Add(result.StringValue);
            }
            Assert.Equal(["Value"], values);

            // Reset for the next iteration.
            reader.Reset();
        }
    }

    [Fact]
    public void GetReader_InvokeNullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var factory = new AllColumnNamesReaderFactory();
        IReadCellResultEnumerator? result = null;
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetCellsReader(null!));
        Assert.Null(result);
    }

    [Fact]
    public void GetReader_InvokeSheetWithoutHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new AllColumnNamesReaderFactory();
        IEnumerable<ReadCellResult>? result = null;
        Assert.Throws<ExcelMappingException>(() => factory.GetCellsReader(sheet));
        Assert.Null(result);
    }

    [Fact]
    public void GetReader_InvokeSheetWithoutHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new AllColumnNamesReaderFactory();
        IEnumerable<ReadCellResult>? result = null;
        Assert.Throws<ExcelMappingException>(() => factory.GetCellsReader(sheet));
        Assert.Null(result);
    }

#pragma warning disable CS0184 // The is operator is being used to test interface implementation
    [Fact]
    public void Interfaces_IColumnNameProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new AllColumnNamesReaderFactory();
        Assert.False(factory is IColumnNameProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndexProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new AllColumnNamesReaderFactory();
        Assert.False(factory is IColumnIndexProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnNamesProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new AllColumnNamesReaderFactory();
        Assert.False(factory is IColumnNamesProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndicesProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new AllColumnNamesReaderFactory();
        Assert.False(factory is IColumnIndicesProviderCellReaderFactory);
    }
#pragma warning restore CS0184
}
