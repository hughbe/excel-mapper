using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelMapper.Mappings.Readers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class EnumerablePropertyMappingTests
    {
        [Fact]
        public void Ctor_MemberInfo_EmptyValueStategy()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);
            Assert.Same(propertyInfo, mapping.Member);

            Assert.NotNull(mapping.ElementMapping);
        }

        [Fact]
        public void WithElementMapping_ValidMapping_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var elementMapping = new SinglePropertyMapping<string>(propertyInfo);

            var mapping = new SubPropertyMapping(propertyInfo);
            Assert.Same(mapping, mapping.WithElementMapping(e =>
            {
                Assert.Same(e, mapping.ElementMapping);
                return elementMapping;
            }));
            Assert.Same(elementMapping, mapping.ElementMapping);
        }

        [Fact]
        public void WithElementMapping_NullMapping_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);

            Assert.Throws<ArgumentNullException>("elementMapping", () => mapping.WithElementMapping(null));
        }

        [Fact]
        public void WithElementMapping_MappingReturnsNull_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);

            Assert.Throws<ArgumentNullException>("elementMapping", () => mapping.WithElementMapping(e => null));
        }

        [Fact]
        public void WithColumnName_SplitValidColumnName_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);
            Assert.Same(mapping, mapping.WithColumnName("ColumnName"));

            SplitColumnReader reader = Assert.IsType<SplitColumnReader>(mapping.ColumnsReader);
            ColumnNameReader innerReader = Assert.IsType<ColumnNameReader>(reader.ColumnReader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_MultiValidColumnName_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnName");
            Assert.Same(mapping, mapping.WithColumnName("ColumnName"));

            SplitColumnReader reader = Assert.IsType<SplitColumnReader>(mapping.ColumnsReader);
            ColumnNameReader innerReader = Assert.IsType<ColumnNameReader>(reader.ColumnReader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);

            Assert.Throws<ArgumentNullException>("columnName", () => mapping.WithColumnName(null));
        }

        [Fact]
        public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);

            Assert.Throws<ArgumentException>("columnName", () => mapping.WithColumnName(string.Empty));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_SplitColumnIndex_Success(int columnIndex)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);
            Assert.Same(mapping, mapping.WithColumnIndex(columnIndex));

            SplitColumnReader reader = Assert.IsType<SplitColumnReader>(mapping.ColumnsReader);
            ColumnIndexReader innerReader = Assert.IsType<ColumnIndexReader>(reader.ColumnReader);
            Assert.Equal(columnIndex, innerReader.ColumnIndex);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_MultiColumnIndex_Success(int columnIndex)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnName");
            Assert.Same(mapping, mapping.WithColumnIndex(columnIndex));

            SplitColumnReader reader = Assert.IsType<SplitColumnReader>(mapping.ColumnsReader);
            ColumnIndexReader innerReader = Assert.IsType<ColumnIndexReader>(reader.ColumnReader);
            Assert.Equal(columnIndex, innerReader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);

            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => mapping.WithColumnIndex(-1));
        }

        public static IEnumerable<object[]> Separators_TestData()
        {
            yield return new object[] { new char[] { ',' } };
            yield return new object[] { new char[] { ';', '-' } };
            yield return new object[] { new List<char> { ';', '-' } };
        }

        [Theory]
        [MemberData(nameof(Separators_TestData))]
        public void WithSeparators_ParamsString_Success(IEnumerable<char> separators)
        {
            char[] separatorsArray = separators.ToArray();

            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);
            Assert.Same(mapping, mapping.WithSeparators(separatorsArray));

            SplitColumnReader reader = Assert.IsType<SplitColumnReader>(mapping.ColumnsReader);
            Assert.Same(separatorsArray, reader.Separators);
        }

        [Theory]
        [MemberData(nameof(Separators_TestData))]
        public void WithSeparators_IEnumerableString_Success(IEnumerable<char> separators)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);
            Assert.Same(mapping, mapping.WithSeparators(separators));

            SplitColumnReader reader = Assert.IsType<SplitColumnReader>(mapping.ColumnsReader);
            Assert.Equal(separators, reader.Separators);
        }

        [Fact]
        public void WithSeparators_NullSeparators_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);

            Assert.Throws<ArgumentNullException>("value", () => mapping.WithSeparators(null));
            Assert.Throws<ArgumentNullException>("value", () => mapping.WithSeparators((IEnumerable<char>)null));
        }

        [Fact]
        public void WithSeparators_EmptySeparators_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo);

            Assert.Throws<ArgumentException>("value", () => mapping.WithSeparators(new char[0]));
            Assert.Throws<ArgumentException>("value", () => mapping.WithSeparators(new List<char>()));
        }

        [Fact]
        public void WithSeperators_MultiMap_ThrowsExcelMappingException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ExcelMappingException>(() => mapping.WithSeparators(new char[0]));
            Assert.Throws<ExcelMappingException>(() => mapping.WithSeparators(new List<char>()));
        }

        [Fact]
        public void WithColumnNames_ParamsString_Success()
        {
            var columnNames = new string[] { "ColumnName1", "ColumnName2", };
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");
            Assert.Same(mapping, mapping.WithColumnNames(columnNames));

            MultipleColumnNamesReader reader = Assert.IsType<MultipleColumnNamesReader>(mapping.ColumnsReader);
            Assert.Same(columnNames, reader.ColumnNames);
        }

        [Fact]
        public void WithColumnNames_IEnumerableString_Success()
        {
            var columnNames = new List<string> { "ColumnName1", "ColumnName2", };
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");
            Assert.Same(mapping, mapping.WithColumnNames(columnNames));

            MultipleColumnNamesReader reader = Assert.IsType<MultipleColumnNamesReader>(mapping.ColumnsReader);
            Assert.Equal(columnNames, reader.ColumnNames);
        }

        [Fact]
        public void WithColumnNames_NullColumnNames_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentNullException>("columnNames", () => mapping.WithColumnNames(null));
            Assert.Throws<ArgumentNullException>("columnNames", () => mapping.WithColumnNames((IEnumerable<string>)null));
        }

        [Fact]
        public void WithColumnNames_EmptyColumnNames_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnNames", () => mapping.WithColumnNames(new string[0]));
            Assert.Throws<ArgumentException>("columnNames", () => mapping.WithColumnNames(new List<string>()));
        }

        [Fact]
        public void WithColumnNames_NullValueInColumnNames_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnNames", () => mapping.WithColumnNames(new string[] { null }));
            Assert.Throws<ArgumentException>("columnNames", () => mapping.WithColumnNames(new List<string> { null }));
        }

        [Fact]
        public void WithColumnIndices_ParamsInt_Success()
        {
            var columnIndices = new int[] { 0, 1 };
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");
            Assert.Same(mapping, mapping.WithColumnIndices(columnIndices));

            MultipleColumnIndicesReader reader = Assert.IsType<MultipleColumnIndicesReader>(mapping.ColumnsReader);
            Assert.Same(columnIndices, reader.ColumnIndices);
        }

        [Fact]
        public void WithColumnIndices_IEnumerableInt_Success()
        {
            var columnIndices = new List<int> { 0, 1 };
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");
            Assert.Same(mapping, mapping.WithColumnIndices(columnIndices));

            MultipleColumnIndicesReader reader = Assert.IsType<MultipleColumnIndicesReader>(mapping.ColumnsReader);
            Assert.Equal(columnIndices, reader.ColumnIndices);
        }

        [Fact]
        public void WithColumnIndices_NullColumnIndices_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentNullException>("columnIndices", () => mapping.WithColumnIndices(null));
            Assert.Throws<ArgumentNullException>("columnIndices", () => mapping.WithColumnIndices((IEnumerable<int>)null));
        }

        [Fact]
        public void WithColumnIndices_EmptyColumnIndices_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnIndices", () => mapping.WithColumnIndices(new int[0]));
            Assert.Throws<ArgumentException>("columnIndices", () => mapping.WithColumnIndices(new List<int>()));
        }

        [Fact]
        public void WithColumnIndices_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var mapping = new SubPropertyMapping(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => mapping.WithColumnIndices(new int[] { -1 }));
            Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => mapping.WithColumnIndices(new List<int> { -1 }));
        }

        private class SubPropertyMapping : EnumerablePropertyMapping<string>
        {
            public SubPropertyMapping(MemberInfo member) : base(member, new SinglePropertyMapping<string>(member))
            {
            }

            public override object CreateFromElements(IEnumerable<string> elements)
            {
                throw new NotImplementedException();
            }
        }

        private class TestClass
        {
            public string[] Value { get; set; }
        }
    }
}
