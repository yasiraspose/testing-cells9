using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class XlsxToPdfConverter
    {
        public static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Desired output PDF file path
            string outputPath = "output.pdf";

            // Convert the Excel workbook to PDF while preserving formatting,
            // layout, and pagination using Aspose.Cells ConversionUtility.
            ConversionUtility.Convert(sourcePath, outputPath);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}