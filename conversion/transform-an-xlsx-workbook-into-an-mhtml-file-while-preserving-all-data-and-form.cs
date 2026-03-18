using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsMhtmlConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Desired output path for the MHTML file
            string destPath = "output.mht";

            // Load the workbook from the XLSX file (uses Workbook(string) constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as MHTML, preserving all data and formatting
            // (uses Workbook.Save(string, SaveFormat) overload)
            workbook.Save(destPath, SaveFormat.MHtml);

            Console.WriteLine($"Workbook successfully converted to MHTML: {destPath}");
        }
    }
}