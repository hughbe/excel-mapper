using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Abstractions;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class AllColumnNamesValueReaderTests
{
    [Fact]
    public void GetReaders_InvokeCanRead_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new AllColumnNamesReaderFactory();
        var reader = factory.GetReader(sheet);
        Assert.NotNull(reader);
        IEnumerable<ReadCellResult>? result = null;
        Assert.True(reader.TryGetValues(importer.Reader, out result));
        Assert.Equal(["Value"], result.Select(r => r.StringValue));
    }

    [Fact]
    public void GetReader_InvokeNullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var factory = new AllColumnNamesReaderFactory();
        IEnumerable<ReadCellResult>? result = null;
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetReader(null!));
        Assert.Null(result);
    }

    [Fact]
    public void GetReader_InvokeSheetWithoutHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new AllColumnNamesReaderFactory();
        IEnumerable<ReadCellResult>? result = null;
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(result);
    }

    [Fact]
    public void GetReader_InvokeSheetWithoutHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new AllColumnNamesReaderFactory();
        IEnumerable<ReadCellResult>? result = null;
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(result);
    }
}
