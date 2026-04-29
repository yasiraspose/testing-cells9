using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XlsxToOdsConversion
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.ods";

            Workbook workbook = new Workbook(sourcePath);
            workbook.Save(destPath, SaveFormat.Ods);

            Console.WriteLine($"Conversion completed successfully: {sourcePath} -> {destPath}");
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}