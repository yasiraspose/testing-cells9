using System;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options to preserve structure, cell types and hierarchy
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if there is only one worksheet
                AlwaysExportAsJsonObject = true,

                // Preserve parent‑child hierarchy (nested structure) in the JSON output
                ExportNestedStructure = true,

                // Convert the Excel file to its JSON struct representation
                ToExcelStruct = true,

                // Include empty cells as null values to keep cell positions
                ExportEmptyCells = true,

                // Do not skip empty rows so that the original layout is retained
                SkipEmptyRows = false
            };

            // Path for the resulting JSON file
            string jsonOutputPath = "output.json";

            // Save the workbook as JSON using the configured options
            workbook.Save(jsonOutputPath, jsonOptions);

            Console.WriteLine($"Workbook '{sourcePath}' has been exported to JSON at '{jsonOutputPath}'.");
        }
    }
}