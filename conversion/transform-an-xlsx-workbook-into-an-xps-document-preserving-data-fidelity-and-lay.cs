using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class XlsxToXpsConverter
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Desired output XPS file path
        string outputPath = "output.xps";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Create XPS save options (default constructor)
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Configure options to preserve layout (one page per sheet disabled)
        xpsOptions.OnePagePerSheet = false;
        xpsOptions.AllColumnsInOnePagePerSheet = false;

        // Save the workbook as XPS using the configured options
        workbook.Save(outputPath, xpsOptions);

        Console.WriteLine("Conversion completed: " + outputPath);
    }
}