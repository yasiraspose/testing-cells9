using System;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    public static class XlsxToHtmlConverter
    {
        /// <summary>
        /// Converts an XLSX workbook to an HTML file while preserving formatting and content.
        /// </summary>
        /// <param name="sourcePath">Full path of the source .xlsx file.</param>
        /// <param name="htmlPath">Full path where the resulting .html file will be saved.</param>
        public static void Convert(string sourcePath, string htmlPath)
        {
            // Load the workbook from the specified XLSX file.
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options to control the conversion.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Preserve workbook and worksheet properties (default true, set explicitly for clarity).
            saveOptions.ExportWorkbookProperties = true;
            saveOptions.ExportWorksheetProperties = true;

            // Optionally export grid lines to match the Excel view.
            saveOptions.ExportGridLines = true;

            // Save the workbook as HTML using the configured options.
            workbook.Save(htmlPath, saveOptions);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsConversion <source.xlsx> <output.html>");
                return;
            }

            string sourcePath = args[0];
            string htmlPath = args[1];

            try
            {
                XlsxToHtmlConverter.Convert(sourcePath, htmlPath);
                Console.WriteLine($"Conversion successful. HTML saved to: {htmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}