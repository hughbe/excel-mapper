using System;
using System.Collections.Generic;
using ExcelMapper;

public class TestClass
{
    public Dictionary<int, string> Value { get; set; } = new Dictionary<int, string>();
}

class Program
{
    static void Main()
    {
        Console.WriteLine("Testing Dictionary<int, string> mapping...");
        
        try
        {
            var map = new ExcelClassMap<TestClass>();
            map.Map(p => p.Value[0]).WithColumnName("Column2");
            Console.WriteLine("Map creation successful!");
            Console.WriteLine($"Map type: {map.GetType()}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine($"Type: {ex.GetType()}");
            Console.WriteLine($"Stack: {ex.StackTrace}");
        }
    }
}