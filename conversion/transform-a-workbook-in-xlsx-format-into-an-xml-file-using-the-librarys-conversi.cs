using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class WorkbookToXmlConverter
    {
        public static void Run()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Desired output XML file path
            string destinationPath = "output.xml";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as XML Spreadsheet format
            workbook.Save(destinationPath, SaveFormat.Xml);

            Console.WriteLine($"Workbook '{sourcePath}' has been successfully converted to XML at '{destinationPath}'.");
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}