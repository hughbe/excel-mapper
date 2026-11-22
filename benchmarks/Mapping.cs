using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelDataReader;
using ExcelMapper.Tests;

namespace ExcelMapper.Benchmarks;

/// <summary>
/// Measures the cost of mapping rows, separately from parsing the file. CSV data is held in memory
/// so that ExcelDataReader's work is small and the same for every benchmark.
/// </summary>
/// <remarks>
/// Built with <c>EXCELMAPPER_VERSION=3.0.3</c>, the benchmarks run against that package instead of
/// <c>src/</c>, to compare versions; the ones for 4.0 features are left out (see README.md).
/// </remarks>
[MemoryDiagnoser]
public class Mapping
{
    private const int RowCount = 50_000;

    private byte[] _validCsv = default!;
    private byte[] _invalidCsv = default!;
    private byte[] _listCsv = default!;
    private byte[] _numbersCsv = default!;
    private byte[] _smallCsv = default!;
    private byte[] _typedNumbersXlsx = default!;
    private byte[] _typedDatesXlsx = default!;
    private byte[] _blankLinesCsv = default!;

    [GlobalSetup]
    public void Setup()
    {
        // ExcelDataReader needs code page encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var valid = new StringBuilder("Name,Age,Score,Salary,Joined,Active,Id,Status,Manager,Notes\n");
        var invalid = new StringBuilder("Name,Age\n");
        var list = new StringBuilder("Name,Values\n");
        var numbers = new StringBuilder("A,B,C,D,E\n");
        for (int i = 0; i < RowCount; i++)
        {
            valid.Append("Name").Append(i).Append(',').Append(i % 90).Append(',').Append(i * 0.5).Append(',')
                .Append(i * 10).Append(',').Append(new DateTime(2024, 1, 2, 3, 4, 5).ToString("G")).Append(',').Append(i % 2 == 0 ? "true" : "false").Append(',')
                .Append(Guid.Empty).Append(',').Append(i % 2 == 0 ? "Active" : "Inactive").Append(',')
                .Append(i % 3 == 0 ? "" : "7").Append(",Some notes\n");
            invalid.Append("Name").Append(i).Append(',').Append(i % 2 == 0 ? "invalid" : "42").Append('\n');
            list.Append("Name").Append(i).Append(",\"").Append(i).Append(',').Append(i + 1).Append(',').Append(i + 2).Append("\"\n");
            numbers.Append(i).Append(',').Append(i + 1).Append(',').Append(i + 2).Append(',').Append(i + 3).Append(',').Append(i + 4).Append('\n');
        }

        _validCsv = Encoding.UTF8.GetBytes(valid.ToString());
        _invalidCsv = Encoding.UTF8.GetBytes(invalid.ToString());
        _listCsv = Encoding.UTF8.GetBytes(list.ToString());
        _numbersCsv = Encoding.UTF8.GetBytes(numbers.ToString());
        _smallCsv = Encoding.UTF8.GetBytes("Name,Age,Score,Status,Joined,Values,Street,Zip,Id,Salary\nA,1,2.5,Active," + new DateTime(2024, 1, 2).ToString("G") + ",\"1,2\",High St,123," + Guid.Empty + ",5\n");

        // Typed number cells, which ExcelDataReader returns as doubles.
        var rows = new object?[RowCount + 1][];
        rows[0] = ["Count", "Total", "Ratio"];
        for (int i = 0; i < RowCount; i++)
        {
            rows[i + 1] = [i, (long)i * 1000, i * 0.25];
        }

        _typedNumbersXlsx = XlsxBuilder.CreateStream(rows).ToArray();

        // Date cells, which ExcelDataReader returns as DateTime.
        var dateRows = new object?[RowCount + 1][];
        dateRows[0] = ["When", "Day"];
        for (int i = 0; i < RowCount; i++)
        {
            dateRows[i + 1] = [new DateTime(2024, 1, 2, 3, 4, 5).AddMinutes(i), new DateTime(2024, 1, 1).AddDays(i % 365)];
        }

        _typedDatesXlsx = XlsxBuilder.CreateStream(dateRows).ToArray();

        // Every other row is blank.
        var blankLines = new StringBuilder("Name,Age\n");
        for (int i = 0; i < RowCount; i++)
        {
            blankLines.Append("Name").Append(i).Append(',').Append(i % 90).Append("\n,\n");
        }

        _blankLinesCsv = Encoding.UTF8.GetBytes(blankLines.ToString());
    }

    private static ExcelImporter CreateImporter(byte[] csv)
        => new(new MemoryStream(csv), ExcelImporterFileType.Csv);

    public enum Status
    {
        Active,
        Inactive
    }

    public class Employee
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public double Score { get; set; }
        public decimal Salary { get; set; }
        public DateTime Joined { get; set; }
        public bool Active { get; set; }
        public Guid Id { get; set; }
        public Status Status { get; set; }
        public int? Manager { get; set; }
        public string Notes { get; set; } = "";
    }

    [Benchmark(Baseline = true)]
    public int ReaderOnly()
    {
        // Reading every cell with ExcelDataReader directly: the floor for any mapping.
        using var reader = ExcelReaderFactory.CreateCsvReader(new MemoryStream(_validCsv));
        int count = 0;
        while (reader.Read())
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetValue(i) is not null)
                {
                    count++;
                }
            }
        }

        return count;
    }

    [Benchmark]
    public int AutoMappedClass()
    {
        using var importer = CreateImporter(_validCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<Employee>())
        {
            count += row.Age;
        }

        return count;
    }

    public struct EmployeeStruct
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public double Score { get; set; }
        public decimal Salary { get; set; }
        public DateTime Joined { get; set; }
        public bool Active { get; set; }
        public Guid Id { get; set; }
        public Status Status { get; set; }
        public int? Manager { get; set; }
        public string Notes { get; set; }
    }

    [Benchmark]
    public int AutoMappedStruct()
    {
        // AutoMappedClass with each row a struct.
        using var importer = CreateImporter(_validCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<EmployeeStruct>())
        {
            count += row.Age;
        }

        return count;
    }

    public class FluentEmployee
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public double Score { get; set; }
        public int? Manager { get; set; }
        public Status Status { get; set; }
        public string Notes { get; set; } = "";
    }

    [Benchmark]
    public int FluentClassMap()
    {
        using var importer = CreateImporter(_validCsv);
        importer.Configuration.RegisterClassMap<FluentEmployee>(c =>
        {
            c.Map(e => e.Name).WithTrim().WithEmptyFallback("none");
            c.Map(e => e.Age).WithInvalidFallback(-1);
            c.Map(e => e.Score).WithConverter(value => double.Parse(value!, CultureInfo.CurrentCulture));
            c.Map(e => e.Manager).WithEmptyFallback(null);
            c.Map(e => e.Status).WithMapping(new Dictionary<string, Status> { { "Active", Status.Active } });
            c.Map(e => e.Notes).WithColumnName("Notes");
        });

        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<FluentEmployee>())
        {
            count += row.Age;
        }

        return count;
    }

    public class InvalidAge
    {
        public string Name { get; set; } = "";

        [ExcelInvalidValue(-1)]
        public int Age { get; set; }
    }

    [Benchmark]
    public int InvalidValuesWithFallback()
    {
        // Half the rows have an invalid age, which gets the fallback value.
        using var importer = CreateImporter(_invalidCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<InvalidAge>())
        {
            count += row.Age;
        }

        return count;
    }

    public class Numbers
    {
        public int Count { get; set; }
        public long Total { get; set; }
        public double Ratio { get; set; }
    }

    [Benchmark]
    public long TypedNumbers()
    {
        using var importer = new ExcelImporter(new MemoryStream(_typedNumbersXlsx));
        long sum = 0;
        foreach (var row in importer.ReadSheet().ReadRows<Numbers>())
        {
            sum += row.Count;
        }

        return sum;
    }

    public class SplitList
    {
        public string Name { get; set; } = "";
        public List<int> Values { get; set; } = default!;
    }

    [Benchmark]
    public int SplitCollection()
    {
        // A list of numbers split from each cell, such as "1,2,3".
        using var importer = CreateImporter(_listCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<SplitList>())
        {
            count += row.Values.Count;
        }

        return count;
    }

    [Benchmark]
    public int DictionaryOfColumns()
    {
        // Every column of each row, as a dictionary.
        using var importer = CreateImporter(_numbersCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<Dictionary<string, int>>())
        {
            count += row.Count;
        }

        return count;
    }

    public class Address
    {
        public string Street { get; set; } = "";
        public int Zip { get; set; }
    }

    public class Wide
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public double? Score { get; set; }
        public Status Status { get; set; }
        public DateTime Joined { get; set; }
        public List<int> Values { get; set; } = default!;
        public Address Home { get; set; } = default!;
        public Guid Id { get; set; }
        public decimal Salary { get; set; }
    }

    [Benchmark]
    public int AutoMapping()
    {
        // The cost of auto-mapping a type, which each importer does the first time it reads it:
        // one row of a type with collection, nested, enum and nullable members.
        using var importer = CreateImporter(_smallCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<Wide>())
        {
            count += row.Age;
        }

        return count;
    }

    public class Dates
    {
        public DateTime When { get; set; }
        public DateTime Day { get; set; }
    }

    [Benchmark]
    public int TypedDates()
    {
        using var importer = new ExcelImporter(new MemoryStream(_typedDatesXlsx));
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<Dates>())
        {
            count += row.Day.Day;
        }

        return count;
    }

    public class NameAndAge
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    [Benchmark]
    public int SkipBlankLines()
    {
        // Half the rows are blank, and skipped.
        using var importer = CreateImporter(_blankLinesCsv);
        importer.Configuration.SkipBlankLines = true;
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<NameAndAge>())
        {
            count += row.Age;
        }

        return count;
    }

    public class WithMissingColumn
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }

        [ExcelOptional]
        public string? Department { get; set; }
    }

    [Benchmark]
    public int OptionalMissingColumn()
    {
        // An optional member whose column the sheet doesn't have.
        using var importer = CreateImporter(_validCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<WithMissingColumn>())
        {
            count += row.Age;
        }

        return count;
    }

#if !EXCELMAPPER_3
    [Benchmark]
    public int AutoMappedClassWithCulture()
    {
        // AutoMappedClass with ExcelImporterConfiguration.Culture set.
        using var importer = CreateImporter(_validCsv);
        importer.Configuration.Culture = CultureInfo.InvariantCulture;
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<Employee>())
        {
            count += row.Age;
        }

        return count;
    }

    public record EmployeeRecord(string Name, int Age, double Score, decimal Salary, DateTime Joined, bool Active, Guid Id, Status Status, int? Manager, string Notes);

    [Benchmark]
    public int AutoMappedRecord()
    {
        using var importer = CreateImporter(_validCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<EmployeeRecord>())
        {
            count += row.Age;
        }

        return count;
    }

    [Benchmark]
    public int RowResultsClass()
    {
        using var importer = CreateImporter(_validCsv);
        int count = 0;
        foreach (var result in importer.ReadSheet().ReadRowResults<Employee>())
        {
            count += result.Value!.Age;
        }

        return count;
    }

    public class InvalidAgeNoFallback
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    [Benchmark]
    public int RowResultsInvalidValues()
    {
        // Half the rows are invalid. Without a fallback, ReadRows would throw at the first.
        using var importer = CreateImporter(_invalidCsv);
        int count = 0;
        foreach (var result in importer.ReadSheet().ReadRowResults<InvalidAgeNoFallback>())
        {
            count += result.IsValid ? result.Value!.Age : 0;
        }

        return count;
    }

    public class ConvertedAge
    {
        public string Name { get; set; } = "";

        [ExcelConverter(typeof(AgeConverter))]
        public int Age { get; set; }
    }

    public class AgeConverter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context)
            => int.TryParse(value.Text, out var age) ? CellConversionResult<int>.Success(age) : CellConversionResult<int>.Success(-1);
    }

    [Benchmark]
    public int TypedConverter()
    {
        // An ICellConverter<int>, which converts without boxing.
        using var importer = CreateImporter(_invalidCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<ConvertedAge>())
        {
            count += row.Age;
        }

        return count;
    }

    [Benchmark]
    public int Region()
    {
        // The same rows as AutoMappedClass, read as a region of the sheet.
        using var importer = CreateImporter(_validCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().Region("A1:J50001").ReadRows<Employee>())
        {
            count += row.Age;
        }

        return count;
    }
#endif
}
