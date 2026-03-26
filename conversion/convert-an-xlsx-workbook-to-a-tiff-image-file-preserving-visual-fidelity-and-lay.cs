using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsConversion
{
    class WorkbookToTiff
    {
        static void Main(string[] args)
        {
            // Input Excel file path (XLSX)
            string inputPath = "input.xlsx";

            // Output TIFF image file path
            string outputPath = "output.tiff";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure image rendering options for TIFF output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,                     // Set output format to TIFF
                TiffCompression = TiffCompression.CompressionLZW, // Use LZW compression for better quality
                HorizontalResolution = 300,                    // Set horizontal DPI
                VerticalResolution = 300,                      // Set vertical DPI
                OnePagePerSheet = true                         // Render each sheet as a separate page
            };

            // Create a renderer for the entire workbook
            WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);

            // Render the whole workbook to a multi‑page TIFF file
            renderer.ToImage(outputPath);

            Console.WriteLine($"Workbook successfully converted to TIFF: {outputPath}");
        }
    }
}