using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversion
{
    public static class ExcelToHtmlConverter
    {
        /// <summary>
        /// Converts an Excel workbook (XLSX) to an HTML file while preserving formatting and content.
        /// </summary>
        /// <param name="excelPath">Full path to the source .xlsx file.</param>
        /// <param name="htmlPath">Full path where the resulting .html file will be saved.</param>
        public static void Convert(string excelPath, string htmlPath)
        {
            // Load the workbook from the specified Excel file.
            Workbook workbook = new Workbook(excelPath);

            // Configure HTML save options to retain formatting.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export grid lines to match the Excel view.
                ExportGridLines = true,

                // Export images as Base64 so that the HTML is self‑contained.
                ExportImagesAsBase64 = true,

                // Preserve formulas (optional, can be set to false if only values are needed).
                ExportFormula = true,

                // Export workbook properties (author, title, etc.).
                ExportWorkbookProperties = true,

                // Export worksheet properties (page setup, etc.).
                ExportWorksheetProperties = true,

                // Use HTML5 standard for better compatibility.
                HtmlVersion = HtmlVersion.Html5
            };

            // Save the workbook as HTML using the configured options.
            workbook.Save(htmlPath, htmlOptions);
        }

        // Example usage
        public static void Main()
        {
            string sourceExcel = @"C:\Data\SampleWorkbook.xlsx";
            string targetHtml = @"C:\Data\SampleWorkbook.html";

            try
            {
                Convert(sourceExcel, targetHtml);
                Console.WriteLine($"Conversion completed successfully. HTML saved to: {targetHtml}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}