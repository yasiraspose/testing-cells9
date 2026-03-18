using System;
using Aspose.Cells;

namespace AsposeCellsMemoryOptimizedDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the large Excel file to be loaded
            string inputPath = "LargeDataFile.xlsx";

            // Create LoadOptions and set memory mode to MemoryPreference
            // This keeps the cells data in a compact format to reduce memory consumption
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            loadOptions.MemorySetting = MemorySetting.MemoryPreference;

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example operation: display the number of worksheets loaded
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");

            // Save the workbook (optional, can be saved to a different format or location)
            string outputPath = "OptimizedOutput.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}' with memory optimization.");
        }
    }
}