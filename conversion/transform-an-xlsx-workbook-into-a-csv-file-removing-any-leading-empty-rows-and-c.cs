using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired CSV output path
        string destPath = "output.csv";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Configure CSV save options to trim leading empty rows and columns
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            TrimLeadingBlankRowAndColumn = true
        };

        // Save the workbook as CSV using the configured options
        workbook.Save(destPath, saveOptions);
    }
}