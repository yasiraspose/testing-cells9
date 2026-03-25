using System;
using Aspose.Cells;

class ConvertXlsxToMhtml
{
    static void Main()
    {
        // Source Excel file
        string sourcePath = "input.xlsx";

        // Destination MHTML file
        string destPath = "output.mht";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as MHTML, preserving all data and formatting
        workbook.Save(destPath, SaveFormat.MHtml);

        Console.WriteLine("Workbook successfully converted to MHTML.");
    }
}