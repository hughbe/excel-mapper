using System;
using System.Collections.Generic;
using ExcelMapper;
using ExcelMapper.Factories;
using ExcelMapper.Abstractions;

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
            
            // Test the factory directly
            var factory = new DictionaryFactory<int, string>();
            Console.WriteLine("Factory creation successful!");
            
            factory.Begin(1);
            factory.Add(0, "test");
            var result = factory.End();
            
            Console.WriteLine($"Factory test successful! Result type: {result.GetType()}");
            
            if (result is Dictionary<int, string> dict)
            {
                Console.WriteLine($"Dictionary contains {dict.Count} items");
                foreach (var kvp in dict)
                {
                    Console.WriteLine($"Key: {kvp.Key}, Value: {kvp.Value}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine($"Type: {ex.GetType()}");
            Console.WriteLine($"Stack: {ex.StackTrace}");
        }
    }
}