using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (XLSX format)
            string sourcePath = "input.xlsx";

            // Desired output path for the XPS file
            string destPath = "output.xps";

            // Convert the Excel workbook to XPS using Aspose.Cells ConversionUtility
            // This method directly handles loading the source file and saving it in the target format.
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
        }
    }
}