using System;
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
[MemoryDiagnoser]
public class Mapping
{
    private const int RowCount = 50_000;

    private byte[] _validCsv = default!;
    private byte[] _invalidCsv = default!;

    [GlobalSetup]
    public void Setup()
    {
        Helpers.GetCsvImporter("A\n1\n").Dispose(); // Registers code pages.

        var valid = new StringBuilder("Name,Age,Score,Salary,Joined,Active,Id,Status,Manager,Notes\n");
        var invalid = new StringBuilder("Name,Age\n");
        for (int i = 0; i < RowCount; i++)
        {
            valid.Append("Name").Append(i).Append(',').Append(i % 90).Append(',').Append(i * 0.5).Append(',')
                .Append(i * 10).Append(",").Append(new DateTime(2024, 1, 2, 3, 4, 5).ToString("G")).Append(",").Append(i % 2 == 0 ? "true" : "false").Append(',')
                .Append(Guid.Empty).Append(',').Append(i % 2 == 0 ? "Active" : "Inactive").Append(',')
                .Append(i % 3 == 0 ? "" : "7").Append(",Some notes\n");
            invalid.Append("Name").Append(i).Append(',').Append(i % 2 == 0 ? "invalid" : "42").Append('\n');
        }

        _validCsv = Encoding.UTF8.GetBytes(valid.ToString());
        _invalidCsv = Encoding.UTF8.GetBytes(invalid.ToString());
    }

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

    public record EmployeeRecord(string Name, int Age, double Score, decimal Salary, DateTime Joined, bool Active, Guid Id, Status Status, int? Manager, string Notes);

    public class InvalidAge
    {
        public string Name { get; set; } = "";

        [ExcelInvalidValue(-1)]
        public int Age { get; set; }
    }

    private ExcelImporter CreateImporter(byte[] csv)
        => new(new MemoryStream(csv), ExcelImporterFileType.Csv);

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
    public int InvalidValuesWithFallback()
    {
        using var importer = CreateImporter(_invalidCsv);
        int count = 0;
        foreach (var row in importer.ReadSheet().ReadRows<InvalidAge>())
        {
            count += row.Age;
        }

        return count;
    }
}
