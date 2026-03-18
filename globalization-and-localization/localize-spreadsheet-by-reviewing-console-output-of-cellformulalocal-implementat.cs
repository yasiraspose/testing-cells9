using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook from an existing XLSX file
        LoadOptions loadOptions = new LoadOptions();
        // Ensure formulas are parsed on load (default behavior)
        loadOptions.ParsingFormulaOnOpen = true;
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Set the workbook locale to German (de-DE) for demonstration
        workbook.Settings.Region = CountryCode.Germany;

        // Work with the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Iterate through all cells that contain formulas
        foreach (Cell cell in worksheet.Cells)
        {
            if (!string.IsNullOrEmpty(cell.Formula))
            {
                Console.WriteLine($"Cell {cell.Name}:");
                // Standard (English) formula
                Console.WriteLine($"  Standard Formula: {cell.Formula}");
                // Locale‑formatted formula via the FormulaLocal property
                Console.WriteLine($"  Localized Formula (FormulaLocal): {cell.FormulaLocal}");
                // Same result using GetFormula with isLocal = true
                Console.WriteLine($"  GetFormula(isLocal:true): {cell.GetFormula(false, true)}");
                Console.WriteLine();
            }
        }

        // Save the workbook (optional, in case any modifications were made)
        workbook.Save("output.xlsx");
    }
}