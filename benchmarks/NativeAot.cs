using System;
using System.IO;
using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Environments;
using BenchmarkDotNet.Jobs;
using ExcelDataReader;

namespace ExcelMapper.Benchmarks;

/// <summary>
/// Measures mapping with a generated <see cref="ExcelMappingContext"/>, which native AOT apps use,
/// with the JIT and compiled with native AOT: the <c>Mapping</c> benchmarks' 50,000 rows of a
/// 10-member class, as a class, a record and a struct, and reading the file alone.
/// </summary>
[Config(typeof(NativeAotConfig))]
[MemoryDiagnoser]
public class NativeAot
{
    private const int RowCount = 50_000;

    private byte[] _csv = default!;

    // Each benchmark with the JIT (the baseline) and as a native AOT app, on the current .NET.
    private sealed class NativeAotConfig : ManualConfig
    {
        public NativeAotConfig()
        {
            AddJob(Job.Default.WithId("JIT").AsBaseline());
            AddJob(Job.Default.WithRuntime(NativeAotRuntime.GetCurrentVersion()).WithId("NativeAOT"));
            HideColumns("Runtime");
        }
    }

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var csv = new StringBuilder("Name,Age,Score,Salary,Joined,Active,Id,Status,Manager,Notes\n");
        for (int i = 0; i < RowCount; i++)
        {
            csv.Append("Name").Append(i).Append(',').Append(i % 90).Append(',').Append(i * 0.5).Append(',')
                .Append(i * 10).Append(',').Append(new DateTime(2024, 1, 2, 3, 4, 5).ToString("G")).Append(',').Append(i % 2 == 0 ? "true" : "false").Append(',')
                .Append(Guid.Empty).Append(',').Append(i % 2 == 0 ? "Active" : "Inactive").Append(',')
                .Append(i % 3 == 0 ? "" : "7").Append(",Some notes\n");
        }

        _csv = Encoding.UTF8.GetBytes(csv.ToString());
    }

    private ExcelSheet ReadSheet() => new ExcelImporter(new MemoryStream(_csv), ExcelImporterFileType.Csv).ReadSheet();

    [Benchmark]
    public int ReaderOnly()
    {
        // Reading every cell with ExcelDataReader directly: native AOT changes this too.
        using var reader = ExcelReaderFactory.CreateCsvReader(new MemoryStream(_csv));
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
    public int Class()
    {
        int count = 0;
        foreach (var row in ReadSheet().ReadRows(BenchmarkContext.Default.Employee))
        {
            count += row.Age;
        }

        return count;
    }

    [Benchmark]
    public int Record()
    {
        int count = 0;
        foreach (var row in ReadSheet().ReadRows(BenchmarkContext.Default.EmployeeRecord))
        {
            count += row.Age;
        }

        return count;
    }

    [Benchmark]
    public int Struct()
    {
        int count = 0;
        foreach (var row in ReadSheet().ReadRows(BenchmarkContext.Default.EmployeeStruct))
        {
            count += row.Age;
        }

        return count;
    }
}

/// <summary>
/// The rows <see cref="NativeAot"/> maps, registered by the source generator.
/// </summary>
[ExcelMappable(typeof(Mapping.Employee))]
[ExcelMappable(typeof(Mapping.EmployeeRecord))]
[ExcelMappable(typeof(Mapping.EmployeeStruct))]
public partial class BenchmarkContext : ExcelMappingContext
{
}
