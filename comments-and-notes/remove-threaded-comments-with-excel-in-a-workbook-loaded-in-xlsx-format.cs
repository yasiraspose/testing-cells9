using System;
using Aspose.Cells;

class RemoveThreadedComments
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Remove all comments (including threaded comments) from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments(); // Clears both legacy and threaded comments
        }

        // Save the workbook after removal
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}