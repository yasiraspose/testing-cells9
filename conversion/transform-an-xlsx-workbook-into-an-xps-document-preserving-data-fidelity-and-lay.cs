using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Desired path for the output XPS document
        string outputPath = "output.xps";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Create XPS save options to preserve layout and fidelity
        XpsSaveOptions xpsOptions = new XpsSaveOptions
        {
            // Keep each sheet's layout across multiple pages (set to false to allow pagination)
            OnePagePerSheet = false,

            // Ensure fonts are checked for compatibility
            CheckWorkbookDefaultFont = true,
            CheckFontCompatibility = true,

            // Optional: set a default font for any missing fonts
            DefaultFont = "Arial"
        };

        // Save the workbook as an XPS document using the specified options
        workbook.Save(outputPath, xpsOptions);

        Console.WriteLine("Workbook successfully converted to XPS.");
    }
}