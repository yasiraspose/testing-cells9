using System;
using Aspose.Cells;

namespace FormulaLocalDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Scenario 1: Read the localized formula of a cell that has a
            // standard (English) formula.
            // ------------------------------------------------------------
            Cell cellA1 = worksheet.Cells["A1"];
            cellA1.Formula = "=SUM(B1:C1)"; // standard formula
            Console.WriteLine("Standard Formula (A1): " + cellA1.Formula);
            Console.WriteLine("Localized Formula (A1) via FormulaLocal: " + cellA1.FormulaLocal);

            // ------------------------------------------------------------
            // Scenario 2: Set a formula using the localized (German) syntax.
            // ------------------------------------------------------------
            Cell cellA2 = worksheet.Cells["A2"];
            cellA2.FormulaLocal = "=SUMME(B1:C1)"; // German localized function name
            Console.WriteLine("\nAfter setting FormulaLocal on A2:");
            Console.WriteLine("Standard Formula (A2): " + cellA2.Formula);
            Console.WriteLine("Localized Formula (A2): " + cellA2.FormulaLocal);

            // ------------------------------------------------------------
            // Scenario 3: Retrieve formulas with GetFormula, toggling the
            // localization flag.
            // ------------------------------------------------------------
            Console.WriteLine("\nGetFormula (English) from A2: " + cellA2.GetFormula(false, false));
            Console.WriteLine("GetFormula (Localized) from A2: " + cellA2.GetFormula(false, true));

            // ------------------------------------------------------------
            // Scenario 4: Use FormulaParseOptions to input a locale‑dependent
            // formula directly (e.g., date formatting in German).
            // ------------------------------------------------------------
            FormulaParseOptions options = new FormulaParseOptions
            {
                LocaleDependent = true,
                R1C1Style = false
            };
            worksheet.Cells["A3"].SetFormula("=TEXT(TODAY(),\"[$-de-DE]dddd, dd mmmm yyyy\")", options);
            Console.WriteLine("\nA3 localized formula via FormulaParseOptions: " + worksheet.Cells["A3"].FormulaLocal);

            // ------------------------------------------------------------
            // Scenario 5: Apply custom globalization settings for another
            // language (Italian) and use the localized function name.
            // ------------------------------------------------------------
            SettableGlobalizationSettings customSettings = new SettableGlobalizationSettings();
            customSettings.SetLocalFunctionName("SUM", "SOMMA", true); // map SUM -> SOMMA
            workbook.Settings.GlobalizationSettings = customSettings;

            // Populate sample data
            worksheet.Cells["B1"].PutValue(1);
            worksheet.Cells["B2"].PutValue(2);
            worksheet.Cells["B3"].PutValue(3);

            // Use the Italian localized function name
            Cell cellA4 = worksheet.Cells["A4"];
            cellA4.FormulaLocal = "=SOMMA(B1:B3)";
            workbook.CalculateFormula();
            Console.WriteLine("\nResult of Italian SUM (SOMMA) in A4: " + cellA4.Value);

            // ------------------------------------------------------------
            // Save the modified workbook
            // ------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}