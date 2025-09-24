using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ManyToOneDictionaryMapTests
    {
        [Fact]
        public void Ctor_MemberInfo_IMultipleCellValuesReader_IValuePipeline_CreateDictionaryFactory()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory);
            Assert.NotNull(propertyMap.ValuePipeline);
        }

        [Fact]
        public void Ctor_NullCellValuesReader_ThrowsArgumentNullException()
        {
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            Assert.Throws<ArgumentNullException>("cellValuesReader", () => new ManyToOneDictionaryMap<string>(null!, valuePipeline, createDictionaryFactory));
        }

        [Fact]
        public void Ctor_NullPipeline_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            Assert.Throws<ArgumentNullException>("valuePipeline", () => new ManyToOneDictionaryMap<string>(cellValuesReader, null!, createDictionaryFactory));
        }

        [Fact]
        public void Ctor_NullCreateDictionaryFactory_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var valuePipeline = new ValuePipeline<string>();
            Assert.Throws<ArgumentNullException>("createDictionaryFactory", () => new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, null!));
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
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory)
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
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory);
            Assert.Throws<ArgumentNullException>("value", () => propertyMap.CellValuesReader = null!);
        }

        [Fact]
        public void WithValueMap_ValidMap_Success()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory);

            var newValuePipeline = new ValuePipeline<string>();
            Assert.Same(propertyMap, propertyMap.WithValueMap(e =>
            {
                Assert.Same(e, propertyMap.ValuePipeline);
                return newValuePipeline;
            }));
            Assert.Same(newValuePipeline, propertyMap.ValuePipeline);
        }

        [Fact]
        public void WithValueMap_NullMap_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory);

            Assert.Throws<ArgumentNullException>("valueMap", () => propertyMap.WithValueMap(null!));
        }

        [Fact]
        public void WithValueMap_MapReturnsNull_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory);

            Assert.Throws<ArgumentNullException>("valueMap", () => propertyMap.WithValueMap(_ => null!));
        }

        [Fact]
        public void WithColumnNames_ParamsString_Success()
        {
            var columnNames = new string[] { "ColumnName1", "ColumnName2" };
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");
            Assert.Same(propertyMap, propertyMap.WithColumnNames(columnNames));

            MultipleColumnNamesValueReader valueReader = Assert.IsType<MultipleColumnNamesValueReader>(propertyMap.CellValuesReader);
            Assert.Same(columnNames, valueReader.ColumnNames);
        }

        [Fact]
        public void WithColumnNames_IEnumerableString_Success()
        {
            var columnNames = new List<string> { "ColumnName1", "ColumnName2" };
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");
            Assert.Same(propertyMap, propertyMap.WithColumnNames((IEnumerable<string>)columnNames));

            MultipleColumnNamesValueReader valueReader = Assert.IsType<MultipleColumnNamesValueReader>(propertyMap.CellValuesReader);
            Assert.Equal(columnNames, valueReader.ColumnNames);
        }

        [Fact]
        public void WithColumnNames_NullColumnNames_ThrowsArgumentNullException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentNullException>("columnNames", () => propertyMap.WithColumnNames(null!));
            Assert.Throws<ArgumentNullException>("columnNames", () => propertyMap.WithColumnNames((IEnumerable<string>)null!));
        }

        [Fact]
        public void WithColumnNames_EmptyColumnNames_ThrowsArgumentException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new string[0]));
            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new List<string>()));
        }

        [Fact]
        public void WithColumnNames_NullValueInColumnNames_ThrowsArgumentException()
        {
            var cellValuesReader = new MultipleColumnNamesValueReader("Column");
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var propertyMap = new ManyToOneDictionaryMap<string>(cellValuesReader, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");

            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new string[] { null! }));
            Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new List<string> { null! }));
        }

        [Fact]
        public void TryGetValue_InvokeCanRead_Success()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            var reader = new MockReader(() => (true, []));
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var map = new ManyToOneDictionaryMap<string>(reader, valuePipeline, createDictionaryFactory);
            object? result = null;
            Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
            Assert.Empty(Assert.IsType<Dictionary<string, string>>(result));
        }

        [Fact]
        public void TryGetValue_InvokeNullSheet_ThrowsArgumentNullException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            var reader = new MockReader(() => (false, null));
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var map = new ManyToOneDictionaryMap<string>(reader, valuePipeline, createDictionaryFactory);
            MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
            object? result = null;
            Assert.Throws<ArgumentNullException>("sheet", () => map.TryGetValue(null!, 0, importer.Reader, member, out result));
            Assert.Null(result);
        }

        [Fact]
        public void TryGetValue_InvokeSheetWithoutHeadingHasHeading_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();

            var reader = new MockReader(() => (false, null));
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var map = new ManyToOneDictionaryMap<string>(reader, valuePipeline, createDictionaryFactory);
            MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
            object? result = null;
            Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
            Assert.Null(result);
        }

        [Fact]
        public void TryGetValue_InvokeSheetWithoutHeadingHasNoHeading_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            var reader = new MockReader(() => (false, null));
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var map = new ManyToOneDictionaryMap<string>(reader, valuePipeline, createDictionaryFactory);
            MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
            object? result = null;
            Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
            Assert.Null(result);
        }

        [Fact]
        public void TryGetValue_InvokeCantReadPropertyInfo_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            var reader = new MockReader(() => (false, null));
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var map = new ManyToOneDictionaryMap<string>(reader, valuePipeline, createDictionaryFactory);
            MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
            object? result = null;
            Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
            Assert.Null(result);
        }

        [Fact]
        public void TryGetValue_InvokeCantReadFieldInfo_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            var reader = new MockReader(() => (false, null));
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var map = new ManyToOneDictionaryMap<string>(reader, valuePipeline, createDictionaryFactory);
            MemberInfo member = typeof(TestClass).GetField(nameof(TestClass._field))!;
            object? result = null;
            Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
            Assert.Null(result);
        }

        [Fact]
        public void TryGetValue_InvokeCantReadEventInfo_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            var reader = new MockReader(() => (false, null));
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var map = new ManyToOneDictionaryMap<string>(reader, valuePipeline, createDictionaryFactory);
            MemberInfo member = typeof(TestClass).GetEvent(nameof(TestClass.Event))!;
            object? result = null;
            Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
            Assert.Null(result);
        }

        [Fact]
        public void TryGetValue_InvokeCantReadNullMember_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            var reader = new MockReader(() => (false, null));
            var valuePipeline = new ValuePipeline<string>();
            CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
            var map = new ManyToOneDictionaryMap<string>(reader, valuePipeline, createDictionaryFactory);
            object? result = null;
            Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
            Assert.Null(result);
        }

        private class MockReader : IMultipleCellValuesReader
        {
            public MockReader(Func<(bool, IEnumerable<ReadCellValueResult>?)> action)
            {
                Action = action;
            }

            public Func<(bool, IEnumerable<ReadCellValueResult>?)> Action { get; }

            public bool TryGetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, [NotNullWhen(true)] out IEnumerable<ReadCellValueResult>? result)
            {
                (bool ret, IEnumerable<ReadCellValueResult>? res) = Action();
                result = res;
                return ret;
            }
        }

        private class TestClass
        {
            public string Value { get; set; } = default!;
#pragma warning disable 0649
            public string _field = default!;
#pragma warning restore 0649

            public event EventHandler Event { add { } remove { } }
        }
    }
}
