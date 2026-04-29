using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertXlsxToMultipleFormats
    {
        /// <summary>
        /// Loads an existing XLSX workbook and saves it as ODS, SXC, and FODS.
        /// </summary>
        /// <param name="sourceXlsxPath">Full path to the source .xlsx file.</param>
        /// <param name="outputFolder">Folder where the converted files will be written.</param>
        public static void Run(string sourceXlsxPath, string outputFolder)
        {
            if (string.IsNullOrWhiteSpace(sourceXlsxPath) || !File.Exists(sourceXlsxPath))
                throw new FileNotFoundException("Source XLSX file not found.", sourceXlsxPath);

            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new ArgumentException("Output folder must be specified.", nameof(outputFolder));

            Directory.CreateDirectory(outputFolder);

            using (var workbook = new Workbook(sourceXlsxPath))
            {
                string baseName = Path.GetFileNameWithoutExtension(sourceXlsxPath);

                string odsPath = Path.Combine(outputFolder, baseName + ".ods");
                string sxcPath = Path.Combine(outputFolder, baseName + ".sxc");
                string fodsPath = Path.Combine(outputFolder, baseName + ".fods");

                workbook.Save(odsPath, SaveFormat.Ods);
                workbook.Save(sxcPath, SaveFormat.Sxc);
                workbook.Save(fodsPath, SaveFormat.Fods);
            }

            Console.WriteLine("Conversion completed successfully:");
            Console.WriteLine($" - ODS : {Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(sourceXlsxPath) + ".ods")}");
            Console.WriteLine($" - SXC : {Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(sourceXlsxPath) + ".sxc")}");
            Console.WriteLine($" - FODS: {Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(sourceXlsxPath) + ".fods")}");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // args[0] = path to source .xlsx file
            // args[1] = output folder (optional)

            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsExamples <sourceXlsxPath> [outputFolder]");
                return;
            }

            string sourcePath = args[0];
            string outputFolder = args.Length > 1 ? args[1] : Path.Combine(Path.GetDirectoryName(sourcePath) ?? "", "Converted");

            try
            {
                ConvertXlsxToMultipleFormats.Run(sourcePath, outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}