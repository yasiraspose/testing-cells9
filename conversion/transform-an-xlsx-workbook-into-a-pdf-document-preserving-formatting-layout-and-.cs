using System;
using Aspose.Cells; // Namespace for Workbook and SaveFormat

class XlsxToPdfConverter
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired PDF output path
        string pdfPath = "output.pdf";

        // Load the workbook from the XLSX file (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as PDF, preserving formatting, layout, and pagination
        workbook.Save(pdfPath, SaveFormat.Pdf);

        Console.WriteLine("Conversion completed: " + pdfPath);
    }
}