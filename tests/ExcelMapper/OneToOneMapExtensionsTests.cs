using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;
using ExcelMapper.Transformers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class OneToOneMapExtensionsTests : ExcelClassMap<Helpers.TestClass>
    {
        [Fact]
        public void WithColumnName_ValidColumnName_Success()
        {
            OneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));

            ColumnNameValueReader reader = Assert.IsType<ColumnNameValueReader>(propertyMap.CellReader);
            Assert.Equal("ColumnName", reader.ColumnName);
        }

        [Fact]
        public void WithColumnNameMatching_ValidColumnName_Success()
        {
            OneToOneMap<string> propertyMap = Map(t => t.Value).WithColumnNameMatching(e => e == "ColumnName");
            Assert.Same(propertyMap, propertyMap.WithColumnNameMatching(e => e == "ColumnName"));

            Assert.IsType<ColumnNameMatchingValueReader>(propertyMap.CellReader);
        }

        [Fact]
        public void WithColumnName_OptionalColumn_Success()
        {
            OneToOneMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));
            Assert.True(propertyMap.Optional);

            ColumnNameValueReader innerReader = Assert.IsType<ColumnNameValueReader>(propertyMap.CellReader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
        {
            OneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Throws<ArgumentNullException>("columnName", () => propertyMap.WithColumnName(null));
        }

        [Fact]
        public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
        {
            OneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Throws<ArgumentException>("columnName", () => propertyMap.WithColumnName(string.Empty));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_ValidColumnIndex_Success(int columnIndex)
        {
            OneToOneMap<string> propertyMap = Map(t => t.Value);
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(columnIndex));

            ColumnIndexValueReader reader = Assert.IsType<ColumnIndexValueReader>(propertyMap.CellReader);
            Assert.Equal(columnIndex, reader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_OptionalColumn_Success()
        {
            OneToOneMap<string> propertyMap = Map(t => t.Value).MakeOptional();
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(1));
            Assert.True(propertyMap.Optional);

            ColumnIndexValueReader innerReader = Assert.IsType<ColumnIndexValueReader>(propertyMap.CellReader);
            Assert.Equal(1, innerReader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            OneToOneMap<string> propertyMap = Map(t => t.Value);

            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => propertyMap.WithColumnIndex(-1));
        }
    }
}
