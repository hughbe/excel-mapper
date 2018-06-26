namespace ExcelMapper.Mappings.Readers.Tests
{
    using System;

    using Xunit;

    public class ColumnNameMatchingValueReaderTests
    {
        [Fact]
        public void Ctor_ColumnName()
        {
            var reader = new ColumnNameMatchingValueReader(e => e == "ColumnName");
            Assert.NotNull(reader);
        }

        [Fact]
        public void Ctor_NullColumnName_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("predicate", () => new ColumnNameMatchingValueReader(null));
        }
    }
}