using System.Linq;
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
        IEnumerable<ReadCellResult>? result = null;
        Assert.True(reader.TryGetValues(importer.Reader, false, out result));
        Assert.Equal(["Value"], result.Select(r => r.StringValue));
    }

    [Fact]
    public void GetReader_InvokeNullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var factory = new AllColumnNamesReaderFactory();
        IEnumerable<ReadCellResult>? result = null;
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
