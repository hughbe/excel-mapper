using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelMapper.Mappings.Readers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class EnumerableExcelPropertyMapTests
    {
        [Fact]
        public void Ctor_MemberInfo_EmptyValueStategy()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);
            Assert.Same(propertyInfo, propertyMap.Member);

            Assert.NotNull(propertyMap.ElementPipeline);
        }

        [Fact]
        public void Ctor_NullPipeline_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            Assert.Throws<ArgumentNullException>("elementPipeline", () => new SubEnumerableExcelPropertyMap(propertyInfo, null));
        }

        [Fact]
        public void WithElementMap_ValidMap_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var elementMap = new OneToOnePropertyMap<string>(propertyInfo);

            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);
            Assert.Same(propertyMap, propertyMap.WithElementMap(e =>
            {
                Assert.Same(e.Pipeline, propertyMap.ElementPipeline);
                return elementMap;
            }));
            Assert.Same(elementMap.Pipeline, propertyMap.ElementPipeline);
        }

        [Fact]
        public void WithElementMap_NullMap_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("elementMap", () => propertyMap.WithElementMap(null));
        }

        [Fact]
        public void WithElementMap_MapReturnsNull_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("elementMap", () => propertyMap.WithElementMap(e => null));
        }

        [Fact]
        public void WithColumnName_SplitValidColumnName_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.ColumnsReader);
            ColumnNameValueReader innerReader = Assert.IsType<ColumnNameValueReader>(valueReader.CellReader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_MultiValidColumnName_Success()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnName");
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.ColumnsReader);
            ColumnNameValueReader innerReader = Assert.IsType<ColumnNameValueReader>(valueReader.CellReader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("columnName", () => propertyMap.WithColumnName(null));
        }

        [Fact]
        public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);

            Assert.Throws<ArgumentException>("columnName", () => propertyMap.WithColumnName(string.Empty));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_SplitColumnIndex_Success(int columnIndex)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(columnIndex));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.ColumnsReader);
            ColumnIndexValueReader innerReader = Assert.IsType<ColumnIndexValueReader>(valueReader.CellReader);
            Assert.Equal(columnIndex, innerReader.ColumnIndex);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_MultiColumnIndex_Success(int columnIndex)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnName");
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(columnIndex));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.ColumnsReader);
            ColumnIndexValueReader innerReader = Assert.IsType<ColumnIndexValueReader>(valueReader.CellReader);
            Assert.Equal(columnIndex, innerReader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);

            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => propertyMap.WithColumnIndex(-1));
        }

        public static IEnumerable<object[]> Separators_Char_TestData()
        {
            yield return new object[] { new char[] { ',' } };
            yield return new object[] { new char[] { ';', '-' } };
            yield return new object[] { new List<char> { ';', '-' } };
        }

        [Theory]
        [MemberData(nameof(Separators_Char_TestData))]
        public void WithSeparators_ParamsChar_Success(IEnumerable<char> separators)
        {
            char[] separatorsArray = separators.ToArray();

            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);
            Assert.Same(propertyMap, propertyMap.WithSeparators(separatorsArray));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.ColumnsReader);
            Assert.Same(separatorsArray, valueReader.Separators);
        }

        [Theory]
        [MemberData(nameof(Separators_Char_TestData))]
        public void WithSeparators_IEnumerableChar_Success(IEnumerable<char> separators)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);
            Assert.Same(propertyMap, propertyMap.WithSeparators(separators));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.ColumnsReader);
            Assert.Equal(separators, valueReader.Separators);
        }

        public static IEnumerable<object[]> Separators_String_TestData()
        {
            yield return new object[] { new string[] { "," } };
            yield return new object[] { new string[] { ";", "-" } };
            yield return new object[] { new List<string> { ";", "-" } };
        }

        [Theory]
        [MemberData(nameof(Separators_String_TestData))]
        public void WithSeparators_ParamsString_Success(IEnumerable<string> separators)
        {
            string[] separatorsArray = separators.ToArray();

            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);
            Assert.Same(propertyMap, propertyMap.WithSeparators(separatorsArray));

            StringSplitCellValueReader valueReader = Assert.IsType<StringSplitCellValueReader>(propertyMap.ColumnsReader);
            Assert.Same(separatorsArray, valueReader.Separators);
        }

        [Theory]
        [MemberData(nameof(Separators_String_TestData))]
        public void WithSeparators_IEnumerableString_Success(IEnumerable<string> separators)
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);
            Assert.Same(propertyMap, propertyMap.WithSeparators(separators));

            StringSplitCellValueReader valueReader = Assert.IsType<StringSplitCellValueReader>(propertyMap.ColumnsReader);
            Assert.Equal(separators, valueReader.Separators);
        }

        [Fact]
        public void WithSeparators_NullSeparators_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);

            Assert.Throws<ArgumentNullException>("value", () => propertyMap.WithSeparators((char[])null));
            Assert.Throws<ArgumentNullException>("value", () => propertyMap.WithSeparators((IEnumerable<char>)null));
            Assert.Throws<ArgumentNullException>("value", () => propertyMap.WithSeparators((string[])null));
            Assert.Throws<ArgumentNullException>("value", () => propertyMap.WithSeparators((IEnumerable<string>)null));
        }

        [Fact]
        public void WithSeparators_EmptySeparators_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo);

            Assert.Throws<ArgumentException>("value", () => propertyMap.WithSeparators(new char[0]));
            Assert.Throws<ArgumentException>("value", () => propertyMap.WithSeparators(new List<char>()));
            Assert.Throws<ArgumentException>("value", () => propertyMap.WithSeparators(new string[0]));
            Assert.Throws<ArgumentException>("value", () => propertyMap.WithSeparators(new List<string>()));
        }

        [Fact]
        public void WithSeperators_MultiMap_ThrowsExcelMappingException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ExcelMappingException>(() => propertyMap.WithSeparators(new char[0]));
            Assert.Throws<ExcelMappingException>(() => propertyMap.WithSeparators(new List<char>()));
            Assert.Throws<ExcelMappingException>(() => propertyMap.WithSeparators(new string[0]));
            Assert.Throws<ExcelMappingException>(() => propertyMap.WithSeparators(new List<string>()));
        }

        [Fact]
        public void WithColumnNames_ParamsString_Success()
        {
            var columnNames = new string[] { "ColumnName1", "ColumnName2" };
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");
            Assert.Same(propertyMap, propertyMap.WithColumnNames(columnNames));

            MultipleColumnNamesValueReader valueReader = Assert.IsType<MultipleColumnNamesValueReader>(propertyMap.ColumnsReader);
            Assert.Same(columnNames, valueReader.ColumnNames);
        }

        [Fact]
        public void WithColumnNames_IEnumerableString_Success()
        {
            var columnNames = new List<string> { "ColumnName1", "ColumnName2" };
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");
            Assert.Same(propertyMap, propertyMap.WithColumnNames(columnNames));

            MultipleColumnNamesValueReader valueReader = Assert.IsType<MultipleColumnNamesValueReader>(propertyMap.ColumnsReader);
            Assert.Equal(columnNames, valueReader.ColumnNames);
        }

        [Fact]
        public void WithColumnNames_NullColumnNames_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentNullException>("columnNames", () => propertyMap.WithColumnNames(null));
            Assert.Throws<ArgumentNullException>("columnNames", () => propertyMap.WithColumnNames((IEnumerable<string>)null));
        }

        [Fact]
        public void WithColumnNames_EmptyColumnNames_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new string[0]));
            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new List<string>()));
        }

        [Fact]
        public void WithColumnNames_NullValueInColumnNames_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new string[] { null }));
            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new List<string> { null }));
        }

        [Fact]
        public void WithColumnIndices_ParamsInt_Success()
        {
            var columnIndices = new int[] { 0, 1 };
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");
            Assert.Same(propertyMap, propertyMap.WithColumnIndices(columnIndices));

            MultipleColumnIndicesValueReader reader = Assert.IsType<MultipleColumnIndicesValueReader>(propertyMap.ColumnsReader);
            Assert.Same(columnIndices, reader.ColumnIndices);
        }

        [Fact]
        public void WithColumnIndices_IEnumerableInt_Success()
        {
            var columnIndices = new List<int> { 0, 1 };
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");
            Assert.Same(propertyMap, propertyMap.WithColumnIndices(columnIndices));

            MultipleColumnIndicesValueReader reader = Assert.IsType<MultipleColumnIndicesValueReader>(propertyMap.ColumnsReader);
            Assert.Equal(columnIndices, reader.ColumnIndices);
        }

        [Fact]
        public void WithColumnIndices_NullColumnIndices_ThrowsArgumentNullException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentNullException>("columnIndices", () => propertyMap.WithColumnIndices(null));
            Assert.Throws<ArgumentNullException>("columnIndices", () => propertyMap.WithColumnIndices((IEnumerable<int>)null));
        }

        [Fact]
        public void WithColumnIndices_EmptyColumnIndices_ThrowsArgumentException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnIndices", () => propertyMap.WithColumnIndices(new int[0]));
            Assert.Throws<ArgumentException>("columnIndices", () => propertyMap.WithColumnIndices(new List<int>()));
        }

        [Fact]
        public void WithColumnIndices_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
        {
            MemberInfo propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Value));
            var propertyMap = new SubEnumerableExcelPropertyMap(propertyInfo).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => propertyMap.WithColumnIndices(new int[] { -1 }));
            Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => propertyMap.WithColumnIndices(new List<int> { -1 }));
        }

        private class SubEnumerableExcelPropertyMap : EnumerableExcelPropertyMap<string>
        {
            public SubEnumerableExcelPropertyMap(MemberInfo member) : base(member, new ValuePipeline())
            {
            }

            public SubEnumerableExcelPropertyMap(MemberInfo member, ValuePipeline elementPipeline) : base(member, elementPipeline)
            {
            }


            protected override object CreateFromElements(IEnumerable<string> elements)
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
