using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Desired output XML file path
            string destPath = "output.xml";

            // Convert the XLSX workbook to XML using Aspose.Cells ConversionUtility
            // This utilizes the built‑in conversion method and avoids manual load/save logic.
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine($"Workbook '{sourcePath}' has been successfully converted to XML at '{destPath}'.");
        }
    }
}