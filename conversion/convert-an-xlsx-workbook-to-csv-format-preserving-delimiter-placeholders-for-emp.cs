using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    class XlsxToCsv
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Path for the resulting CSV file
            string destinationPath = "output.csv";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Configure CSV (text) save options
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                // Preserve delimiter placeholders for completely blank rows
                KeepSeparatorsForBlankRow = true,

                // Optional: keep leading blank rows/columns to maintain alignment
                TrimLeadingBlankRowAndColumn = false,

                // Use UTF-8 encoding for the CSV output
                Encoding = Encoding.UTF8
            };

            // Save the workbook as CSV using the configured options
            workbook.Save(destinationPath, saveOptions);

            Console.WriteLine($"Workbook successfully converted to CSV at: {destinationPath}");
        }
    }
}