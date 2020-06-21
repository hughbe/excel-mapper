using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ManyToOneEnumerableMapTests
    {
        [Fact]
        public void Ctor_IMultipleCellValuesReader_IValuePipeline_CreateElementsFactory()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);
            Assert.False(propertyMap.Optional);
            Assert.NotNull(propertyMap.ElementPipeline);
        }

        [Fact]
        public void Ctor_NullCellValuesReader_ThrowsArgumentNullException()
        {
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            Assert.Throws<ArgumentNullException>("cellValuesReader", () => new ManyToOneEnumerableMap<string>(null, elementPipeline, createElementsFactory));
        }

        [Fact]
        public void Ctor_NullPipeline_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            Assert.Throws<ArgumentNullException>("elementPipeline", () => new ManyToOneEnumerableMap<string>(cellValuesReader, null, createElementsFactory));
        }

        [Fact]
        public void Ctor_NullCreateElementsFactory_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            Assert.Throws<ArgumentNullException>("createElementsFactory", () => new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, null));
        }

        public static IEnumerable<object[]> CellValuesReader_Set_TestData()
        {
            yield return new object[] { new MultipleColumnNamesValueReader("Column") };
        }

        [Theory]
        [MemberData(nameof(CellValuesReader_Set_TestData))]
        public void CellValuesReader_SetValid_GetReturnsExpected(IMultipleCellValuesReader value)
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory)
            {
                CellValuesReader = value
            };
            Assert.Same(value, propertyMap.CellValuesReader);

            // Set same.
            propertyMap.CellValuesReader = value;
            Assert.Same(value, propertyMap.CellValuesReader);
        }

        [Fact]
        public void CellValuesReader_SetNull_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);
            Assert.Throws<ArgumentNullException>("value", () => propertyMap.CellValuesReader = null);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Optional_Set_GetReturnsExpected(bool value)
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory)
            {
                Optional = value
            };
            Assert.Equal(value, propertyMap.Optional);

            // Set same.
            propertyMap.Optional = value;
            Assert.Equal(value, propertyMap.Optional);

            // Set different.
            propertyMap.Optional = !value;
            Assert.Equal(!value, propertyMap.Optional);
        }

        [Fact]
        public void MakeOptional_HasMapper_ReturnsExpected()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);
            Assert.False(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.MakeOptional());
            Assert.True(propertyMap.Optional);
            Assert.Same(cellValuesReader, propertyMap.CellValuesReader);
        }

        [Fact]
        public void MakeOptional_AlreadyOptional_ReturnsExpected()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);
            Assert.Same(propertyMap, propertyMap.MakeOptional());
            Assert.True(propertyMap.Optional);
            Assert.Same(propertyMap, propertyMap.MakeOptional());
            Assert.True(propertyMap.Optional);
            Assert.Same(cellValuesReader, propertyMap.CellValuesReader);
        }

        [Fact]
        public void WithElementMap_ValidMap_Success()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);

            var newElementPipeline = new ValuePipeline<string>();
            Assert.Same(propertyMap, propertyMap.WithElementMap(e =>
            {
                Assert.Same(e, propertyMap.ElementPipeline);
                return newElementPipeline;
            }));
            Assert.Same(newElementPipeline, propertyMap.ElementPipeline);
        }

        [Fact]
        public void WithElementMap_NullMap_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);

            Assert.Throws<ArgumentNullException>("elementMap", () => propertyMap.WithElementMap(null));
        }

        [Fact]
        public void WithElementMap_MapReturnsNull_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);

            Assert.Throws<ArgumentNullException>("elementMap", () => propertyMap.WithElementMap(e => null));
        }

        [Fact]
        public void WithColumnName_SplitValidColumnName_Success()
        {
            var cellValuesReader = new CharSplitCellValueReader(new ColumnNameValueReader("Column"));
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.CellValuesReader);
            ColumnNameValueReader innerReader = Assert.IsType<ColumnNameValueReader>(valueReader.CellReader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_MultiValidColumnName_Success()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnName");
            Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.CellValuesReader);
            ColumnNameValueReader innerReader = Assert.IsType<ColumnNameValueReader>(valueReader.CellReader);
            Assert.Equal("ColumnName", innerReader.ColumnName);
        }

        [Fact]
        public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);

            Assert.Throws<ArgumentNullException>("columnName", () => propertyMap.WithColumnName(null));
        }

        [Fact]
        public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);

            Assert.Throws<ArgumentException>("columnName", () => propertyMap.WithColumnName(string.Empty));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_SplitColumnIndex_Success(int columnIndex)
        {
            var cellValuesReader = new CharSplitCellValueReader(new ColumnNameValueReader("Column"));
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(columnIndex));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.CellValuesReader);
            ColumnIndexValueReader innerReader = Assert.IsType<ColumnIndexValueReader>(valueReader.CellReader);
            Assert.Equal(columnIndex, innerReader.ColumnIndex);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void WithColumnIndex_MultiColumnIndex_Success(int columnIndex)
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnName");
            Assert.Same(propertyMap, propertyMap.WithColumnIndex(columnIndex));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.CellValuesReader);
            ColumnIndexValueReader innerReader = Assert.IsType<ColumnIndexValueReader>(valueReader.CellReader);
            Assert.Equal(columnIndex, innerReader.ColumnIndex);
        }

        [Fact]
        public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);

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

            var cellValuesReader = new StringSplitCellValueReader(new ColumnNameValueReader("Column"));
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);
            Assert.Same(propertyMap, propertyMap.WithSeparators(separatorsArray));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.CellValuesReader);
            Assert.Same(separatorsArray, valueReader.Separators);
        }

        [Theory]
        [MemberData(nameof(Separators_Char_TestData))]
        public void WithSeparators_IEnumerableChar_Success(IEnumerable<char> separators)
        {
            var cellValuesReader = new StringSplitCellValueReader(new ColumnNameValueReader("Column"));
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);
            Assert.Same(propertyMap, propertyMap.WithSeparators(separators));

            CharSplitCellValueReader valueReader = Assert.IsType<CharSplitCellValueReader>(propertyMap.CellValuesReader);
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

            var cellValuesReader = new StringSplitCellValueReader(new ColumnNameValueReader("Column"));
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);
            Assert.Same(propertyMap, propertyMap.WithSeparators(separatorsArray));

            StringSplitCellValueReader valueReader = Assert.IsType<StringSplitCellValueReader>(propertyMap.CellValuesReader);
            Assert.Same(separatorsArray, valueReader.Separators);
        }

        [Theory]
        [MemberData(nameof(Separators_String_TestData))]
        public void WithSeparators_IEnumerableString_Success(IEnumerable<string> separators)
        {
            var cellValuesReader = new StringSplitCellValueReader(new ColumnNameValueReader("Column"));
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);
            Assert.Same(propertyMap, propertyMap.WithSeparators(separators));

            StringSplitCellValueReader valueReader = Assert.IsType<StringSplitCellValueReader>(propertyMap.CellValuesReader);
            Assert.Equal(separators, valueReader.Separators);
        }

        [Fact]
        public void WithSeparators_NullSeparators_ThrowsArgumentNullException()
        {
            var cellValuesReader = new StringSplitCellValueReader(new ColumnNameValueReader("Column"));
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);

            Assert.Throws<ArgumentNullException>("value", () => propertyMap.WithSeparators((char[])null));
            Assert.Throws<ArgumentNullException>("value", () => propertyMap.WithSeparators((IEnumerable<char>)null));
            Assert.Throws<ArgumentNullException>("value", () => propertyMap.WithSeparators((string[])null));
            Assert.Throws<ArgumentNullException>("value", () => propertyMap.WithSeparators((IEnumerable<string>)null));
        }

        [Fact]
        public void WithSeparators_EmptySeparators_ThrowsArgumentException()
        {
            var cellValuesReader = new StringSplitCellValueReader(new ColumnNameValueReader("Column"));
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory);

            Assert.Throws<ArgumentException>("value", () => propertyMap.WithSeparators(new char[0]));
            Assert.Throws<ArgumentException>("value", () => propertyMap.WithSeparators(new List<char>()));
            Assert.Throws<ArgumentException>("value", () => propertyMap.WithSeparators(new string[0]));
            Assert.Throws<ArgumentException>("value", () => propertyMap.WithSeparators(new List<string>()));
        }

        [Fact]
        public void WithSeperators_MultiMap_ThrowsExcelMappingException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");

            Assert.Throws<ExcelMappingException>(() => propertyMap.WithSeparators(new char[0]));
            Assert.Throws<ExcelMappingException>(() => propertyMap.WithSeparators(new List<char>()));
            Assert.Throws<ExcelMappingException>(() => propertyMap.WithSeparators(new string[0]));
            Assert.Throws<ExcelMappingException>(() => propertyMap.WithSeparators(new List<string>()));
        }

        [Fact]
        public void WithColumnNames_ParamsString_Success()
        {
            var columnNames = new string[] { "ColumnName1", "ColumnName2" };
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");
            Assert.Same(propertyMap, propertyMap.WithColumnNames(columnNames));

            MultipleColumnNamesValueReader valueReader = Assert.IsType<MultipleColumnNamesValueReader>(propertyMap.CellValuesReader);
            Assert.Same(columnNames, valueReader.ColumnNames);
        }

        [Fact]
        public void WithColumnNames_IEnumerableString_Success()
        {
            var columnNames = new List<string> { "ColumnName1", "ColumnName2" };
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");
            Assert.Same(propertyMap, propertyMap.WithColumnNames((IEnumerable<string>)columnNames));

            MultipleColumnNamesValueReader valueReader = Assert.IsType<MultipleColumnNamesValueReader>(propertyMap.CellValuesReader);
            Assert.Equal(columnNames, valueReader.ColumnNames);
        }

        [Fact]
        public void WithColumnNames_NullColumnNames_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentNullException>("columnNames", () => propertyMap.WithColumnNames(null));
            Assert.Throws<ArgumentNullException>("columnNames", () => propertyMap.WithColumnNames((IEnumerable<string>)null));
        }

        [Fact]
        public void WithColumnNames_EmptyColumnNames_ThrowsArgumentException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new string[0]));
            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new List<string>()));
        }

        [Fact]
        public void WithColumnNames_NullValueInColumnNames_ThrowsArgumentException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new string[] { null }));
            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new List<string> { null }));
        }

        [Fact]
        public void WithColumnIndices_ParamsInt_Success()
        {
            var columnIndices = new int[] { 0, 1 };
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");
            Assert.Same(propertyMap, propertyMap.WithColumnIndices(columnIndices));

            MultipleColumnIndicesValueReader reader = Assert.IsType<MultipleColumnIndicesValueReader>(propertyMap.CellValuesReader);
            Assert.Same(columnIndices, reader.ColumnIndices);
        }

        [Fact]
        public void WithColumnIndices_IEnumerableInt_Success()
        {
            var columnIndices = new List<int> { 0, 1 };
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");
            Assert.Same(propertyMap, propertyMap.WithColumnIndices(columnIndices));

            MultipleColumnIndicesValueReader reader = Assert.IsType<MultipleColumnIndicesValueReader>(propertyMap.CellValuesReader);
            Assert.Equal(columnIndices, reader.ColumnIndices);
        }

        [Fact]
        public void WithColumnIndices_NullColumnIndices_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentNullException>("columnIndices", () => propertyMap.WithColumnIndices(null));
            Assert.Throws<ArgumentNullException>("columnIndices", () => propertyMap.WithColumnIndices((IEnumerable<int>)null));
        }

        [Fact]
        public void WithColumnIndices_EmptyColumnIndices_ThrowsArgumentException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnIndices", () => propertyMap.WithColumnIndices(new int[0]));
            Assert.Throws<ArgumentException>("columnIndices", () => propertyMap.WithColumnIndices(new List<int>()));
        }

        [Fact]
        public void WithColumnIndices_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var elementPipeline = new ValuePipeline<string>();
            CreateElementsFactory<string> createElementsFactory = elements => elements;
            var propertyMap = new ManyToOneEnumerableMap<string>(cellValuesReader, elementPipeline, createElementsFactory).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => propertyMap.WithColumnIndices(new int[] { -1 }));
            Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => propertyMap.WithColumnIndices(new List<int> { -1 }));
        }

        private class TestClass
        {
            public string[] Value { get; set; }
        }
    }
}
