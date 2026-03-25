using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class WorkbookConversionDemo
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "source.xlsx";

        // Load the workbook using the constructor that accepts a file path
        Workbook workbook = new Workbook(sourcePath);

        // -----------------------------------------------------------------
        // 1. Convert to CSV using the ConversionUtility (static conversion method)
        // -----------------------------------------------------------------
        string csvPath = "output.csv";
        ConversionUtility.Convert(sourcePath, csvPath);
        Console.WriteLine($"Workbook converted to CSV: {csvPath}");

        // -----------------------------------------------------------------
        // 2. Convert to TSV using the workbook's Save method with SaveFormat.Tsv
        // -----------------------------------------------------------------
        string tsvPath = "output.tsv";
        workbook.Save(tsvPath, SaveFormat.Tsv);
        Console.WriteLine($"Workbook saved as TSV: {tsvPath}");

        // -----------------------------------------------------------------
        // 3. Convert to a generic TXT file (tab‑delimited) using TxtSaveOptions
        // -----------------------------------------------------------------
        string txtPath = "output.txt";

        // Create TxtSaveOptions; specify the separator (tab character) and format (CSV)
        TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
        txtOptions.Separator = '\t';               // Use tab as the delimiter
        txtOptions.ExportAllSheets = true;         // Export all worksheets

        // Save the workbook as a TXT file with the defined options
        workbook.Save(txtPath, txtOptions);
        Console.WriteLine($"Workbook saved as TXT (tab‑delimited): {txtPath}");

        // Clean up
        workbook.Dispose();
    }
}