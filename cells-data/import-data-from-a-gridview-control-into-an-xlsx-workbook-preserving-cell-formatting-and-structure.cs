using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Sample DataTable
        DataTable dt = new DataTable("Sample");
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("Name", typeof(string));
        dt.Rows.Add(1, "Alice");
        dt.Rows.Add(2, "Bob");

        string outputPath = "DataTableExport.xlsx";
        ExportDataTableToExcel(dt, outputPath);
        Console.WriteLine($"Exported to {outputPath}");
    }

    public static void ExportDataTableToExcel(DataTable dataTable, string outputPath)
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        var cells = worksheet.Cells;

        // Write column headers
        for (int col = 0; col < dataTable.Columns.Count; col++)
        {
            cells[0, col].PutValue(dataTable.Columns[col].ColumnName);
        }

        // Write data rows
        for (int row = 0; row < dataTable.Rows.Count; row++)
        {
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                cells[row + 1, col].PutValue(dataTable.Rows[row][col]);
            }
        }

        workbook.Save(outputPath);
    }
}