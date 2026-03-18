using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook locale (e.g., German) to demonstrate localization
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and a target cell (A1)
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Example 1: Set a formula using the standard (English) syntax
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized forms
            Console.WriteLine("Standard Formula : " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Example 2: Set a formula using the localized (German) syntax
            // In German, SUM is SUMME
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // After setting FormulaLocal, the standard Formula property is updated automatically
            Console.WriteLine("\nAfter assigning FormulaLocal:");
            Console.WriteLine("Standard Formula : " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Demonstrate GetFormula with localization flags
            // isR1C1 = false (A1 style), isLocal = true (localized)
            string localized = cell.GetFormula(false, true);
            string standard = cell.GetFormula(false, false);
            Console.WriteLine("\nGetFormula results:");
            Console.WriteLine("Standard (English) : " + standard);
            Console.WriteLine("Localized (German) : " + localized);

            // Save the modified workbook (lifecycle rule: save)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}