using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class ExcelToPptxConverter
    {
        public static void Run()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Desired output PowerPoint file
            string destPath = "output.pptx";

            // Convert the Excel file to PPTX using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelToPptxConverter.Run();
        }
    }
}