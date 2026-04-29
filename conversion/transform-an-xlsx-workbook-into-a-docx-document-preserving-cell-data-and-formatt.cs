using System;
using Aspose.Cells;
using Aspose.Cells.Saving; // Namespace for DocxSaveOptions (if needed)

namespace AsposeCellsConversion
{
    public class XlsxToDocxConverter
    {
        public static void Convert(string sourceXlsxPath, string destinationDocxPath)
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(sourceXlsxPath);

            // Initialize DOCX save options to preserve formatting and enable editable shapes
            DocxSaveOptions docxOptions = new DocxSaveOptions();
            docxOptions.SaveAsEditableShaps = true; // optional: makes shapes editable in Word

            // Save the workbook as a DOCX document using the specified options
            workbook.Save(destinationDocxPath, docxOptions);
        }

        // Example usage
        public static void Main()
        {
            string sourceFile = "sample.xlsx";
            string destFile = "sample.docx";

            Convert(sourceFile, destFile);
            Console.WriteLine($"Conversion completed: '{sourceFile}' -> '{destFile}'");
        }
    }
}