using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class Converter
    {
        /// <summary>
        /// Loads an existing XLSX workbook and saves it to ODS, SXC, and FODS formats.
        /// The Save method preserves all data, formulas, styles, and layout.
        /// </summary>
        /// <param name="sourceFilePath">Full path to the source .xlsx file.</param>
        public static void ConvertXlsxToOdsSxcFods(string sourceFilePath)
        {
            // Load the workbook from the specified XLSX file.
            Workbook workbook = new Workbook(sourceFilePath);

            // Prepare output file names with appropriate extensions.
            string odsPath = System.IO.Path.ChangeExtension(sourceFilePath, ".ods");
            string sxcPath = System.IO.Path.ChangeExtension(sourceFilePath, ".sxc");
            string fodsPath = System.IO.Path.ChangeExtension(sourceFilePath, ".fods");

            // Save the workbook as OpenDocument Spreadsheet (ODS).
            workbook.Save(odsPath, SaveFormat.Ods);

            // Save the workbook as StarOffice Calc Spreadsheet (SXC).
            workbook.Save(sxcPath, SaveFormat.Sxc);

            // Save the workbook as OpenDocument Flat XML Spreadsheet (FODS).
            workbook.Save(fodsPath, SaveFormat.Fods);

            // Release resources.
            workbook.Dispose();

            Console.WriteLine("Conversion completed successfully:");
            Console.WriteLine($"  ODS  -> {odsPath}");
            Console.WriteLine($"  SXC  -> {sxcPath}");
            Console.WriteLine($"  FODS -> {fodsPath}");
        }

        // Entry point for the console application.
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the full path to the source .xlsx file as a command‑line argument.");
                return;
            }

            string sourceFilePath = args[0];

            if (!System.IO.File.Exists(sourceFilePath))
            {
                Console.WriteLine($"File not found: {sourceFilePath}");
                return;
            }

            try
            {
                ConvertXlsxToOdsSxcFods(sourceFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}