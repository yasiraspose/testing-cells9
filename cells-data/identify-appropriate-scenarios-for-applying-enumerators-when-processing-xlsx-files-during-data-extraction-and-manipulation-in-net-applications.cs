using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsEnumeratorScenarios
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and populate it with sample data.
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Simple data for cell enumeration
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);
            sheet.Cells["A4"].PutValue("Charlie");
            sheet.Cells["B4"].PutValue(78);

            // ------------------------------------------------------------
            // Scenario 1: Enumerate all cells in the worksheet.
            // ------------------------------------------------------------
            Console.WriteLine("=== All Cells ===");
            IEnumerator cellEnum = sheet.Cells.GetEnumerator();
            while (cellEnum.MoveNext())
            {
                Cell c = (Cell)cellEnum.Current;
                Console.WriteLine($"{c.Name}: {c.Value}");
            }

            // ------------------------------------------------------------
            // Scenario 2: Enumerate rows in normal order.
            // ------------------------------------------------------------
            Console.WriteLine("\n=== Rows (normal order) ===");
            IEnumerator rowEnum = sheet.Cells.Rows.GetEnumerator();
            while (rowEnum.MoveNext())
            {
                Row r = (Row)rowEnum.Current;
                Console.WriteLine($"Row {r.Index} Height={r.Height}");
            }

            // ------------------------------------------------------------
            // Scenario 3: Enumerate rows in reverse order.
            // ------------------------------------------------------------
            Console.WriteLine("\n=== Rows (reverse order) ===");
            IEnumerator revRowEnum = sheet.Cells.Rows.GetEnumerator(true, false);
            while (revRowEnum.MoveNext())
            {
                Row r = (Row)revRowEnum.Current;
                Console.WriteLine($"Row {r.Index}");
            }

            // ------------------------------------------------------------
            // Scenario 4: Enumerate cells within a specific range.
            // ------------------------------------------------------------
            Console.WriteLine("\n=== Range B2:C4 ===");
            AsposeRange range = sheet.Cells.CreateRange("B2:C4");
            IEnumerator rangeEnum = range.GetEnumerator();
            while (rangeEnum.MoveNext())
            {
                Cell c = (Cell)rangeEnum.Current;
                Console.WriteLine($"{c.Name}: {c.Value}");
            }

            // ------------------------------------------------------------
            // Scenario 5: Enumerate external links.
            // ------------------------------------------------------------
            Console.WriteLine("\n=== External Links ===");
            // Add dummy external links for demonstration
            sheet.Workbook.Worksheets.ExternalLinks.Add("link1.xlam", new string[] { "Sheet1!A1" });
            sheet.Workbook.Worksheets.ExternalLinks.Add("link2.xlam", new string[] { "Sheet1!B2" });
            IEnumerator linkEnum = sheet.Workbook.Worksheets.ExternalLinks.GetEnumerator();
            int linkIdx = 0;
            while (linkEnum.MoveNext())
            {
                ExternalLink link = (ExternalLink)linkEnum.Current;
                Console.WriteLine($"Link {++linkIdx}: {link.DataSource}");
            }

            // ------------------------------------------------------------
            // Scenario 6: Enumerate pivot table fields and items.
            // ------------------------------------------------------------
            // Prepare data for a pivot table
            sheet.Cells["D1"].PutValue("Region");
            sheet.Cells["E1"].PutValue("Sales");
            sheet.Cells["D2"].PutValue("North");
            sheet.Cells["E2"].PutValue(1200);
            sheet.Cells["D3"].PutValue("South");
            sheet.Cells["E3"].PutValue(950);
            sheet.Cells["D4"].PutValue("East");
            sheet.Cells["E4"].PutValue(780);
            sheet.Cells["D5"].PutValue("West");
            sheet.Cells["E5"].PutValue(660);

            // Create the pivot table
            int pivotIdx = sheet.PivotTables.Add("D1:E5", "G3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Region");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // Enumerate row fields
            Console.WriteLine("\n=== Pivot Row Fields ===");
            IEnumerator rowFieldEnum = pivot.RowFields.GetEnumerator();
            while (rowFieldEnum.MoveNext())
            {
                PivotField field = (PivotField)rowFieldEnum.Current;
                Console.WriteLine($"Field: {field.Name}");

                // Enumerate items of each field
                Console.WriteLine("  Items:");
                IEnumerator itemEnum = field.PivotItems.GetEnumerator();
                while (itemEnum.MoveNext())
                {
                    PivotItem item = (PivotItem)itemEnum.Current;
                    Console.WriteLine($"    {item.Value}");
                }
            }

            // ------------------------------------------------------------
            // Save the workbook.
            // ------------------------------------------------------------
            workbook.Save("EnumeratorScenarios.xlsx");
            Console.WriteLine("\nWorkbook saved as 'EnumeratorScenarios.xlsx'.");
        }
    }
}