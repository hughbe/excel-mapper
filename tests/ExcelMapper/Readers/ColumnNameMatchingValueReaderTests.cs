using System;
using Xunit;

namespace ExcelMapper.Readers.Tests
{
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
