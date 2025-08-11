using System;
using ExcelMapper.Utilities;
using Xunit;

namespace ExcelMapper.Tests.Utilities
{
    public class FileUtilsTests
    {
        [Theory]
        [InlineData(".csv", true)]
        [InlineData(".CSV", true)]
        [InlineData(".xls", false)]
        [InlineData("csv", true)]
        [InlineData(null, false)]
        [InlineData("", false)]
        [InlineData("  ", false)]
        [InlineData(" .csv ", true)]
        public void IsCsvExtension_Invoke_ReturnsExpected(string extension, bool expected)
        {
            Assert.Equal(expected, FileUtils.IsCsvExtension(extension));
        }

        [Theory]
        [InlineData(".xls", true)]
        [InlineData(".XLS", true)]
        [InlineData(".xlsx", true)]
        [InlineData(".XLSX", true)]
        [InlineData(".xlsm", true)]
        [InlineData(".XLSM", true)]
        [InlineData(".xlsb", true)]
        [InlineData(".XLSB", true)]
        [InlineData(".csv", false)]
        [InlineData("xls", true)]
        [InlineData(null, false)]
        [InlineData("", false)]
        [InlineData("  ", false)]
        [InlineData(" .xls ", true)]
        public void IsExcelExtension_Invoke_ReturnsExpected(string extension, bool expected)
        {
            Assert.Equal(expected, FileUtils.IsExcelExtension(extension));
        }

        [Theory]
        [InlineData(".csv", true)]
        [InlineData(".xls", true)]
        [InlineData(".xlsx", true)]
        [InlineData(".xlsm", true)]
        [InlineData(".xlsb", true)]
        [InlineData(".CSV", true)]
        [InlineData(".XLS", true)]
        [InlineData(".doc", false)]
        [InlineData(null, false)]
        [InlineData("", false)]
        [InlineData("  ", false)]
        [InlineData(" .csv ", true)]
        public void IsSupportedExtension_Invoke_ReturnsExpected(string extension, bool expected)
        {
            Assert.Equal(expected, FileUtils.IsSupportedExtension(extension));
        }

        [Theory]
        [InlineData(".csv", ".csv")]
        [InlineData(".CSV", ".csv")]
        [InlineData("csv", ".csv")]
        [InlineData("CSV", ".csv")]
        [InlineData("  .cSV  ", ".csv")]
        [InlineData("  CSV  ", ".csv")]
        [InlineData(null, null)]
        [InlineData("", null)]
        [InlineData("  ", null)]
        public void NormalizeExtension_Invoke_ReturnsExpected(string? extension, string? expected)
        {
            Assert.Equal(expected, FileUtils.NormalizeExtension(extension));
        }
    }
}
