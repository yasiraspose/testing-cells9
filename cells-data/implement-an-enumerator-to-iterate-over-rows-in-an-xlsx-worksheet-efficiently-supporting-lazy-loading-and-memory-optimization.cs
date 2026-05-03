using System;
using System.Collections;
using Aspose.Cells;

class RowProcessingHandler : LightCellsDataHandler
{
    // Called when a worksheet starts processing
    public bool StartSheet(Worksheet sheet)
    {
        Console.WriteLine($"Processing sheet: {sheet.Name}");
        return true; // Continue processing this sheet
    }

    // Called before each row is read; return true to process the row
    public bool StartRow(int rowIndex)
    {
        return true; // Process every row
    }

    // Called after a row object is created; can read row properties here
    public bool ProcessRow(Row row)
    {
        Console.WriteLine($"Row {row.Index} Height: {row.Height}");
        return true; // Continue to cells of this row
    }

    // Called before each cell in the current row; return true to process the cell
    public bool StartCell(int columnIndex)
    {
        return true; // Process every cell
    }

    // Called after a cell object is created; can read cell value here
    public bool ProcessCell(Cell cell)
    {
        Console.WriteLine($"  Cell {cell.Name} = {cell.StringValue}");
        return true;
    }
}

class RowEnumeratorDemo
{
    static void Main()
    {
        // Load a large XLSX file in LightCells (streaming) mode for low memory usage
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.LightCellsDataHandler = new RowProcessingHandler();

        // Path to the source workbook (replace with actual file)
        string inputPath = "LargeFile.xlsx";

        // Workbook is loaded row‑by‑row via the handler above
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // After streaming processing, you can still enumerate rows normally if needed
        Worksheet worksheet = workbook.Worksheets[0];

        // Use a synchronized enumerator to safely traverse rows without modifying the collection
        IEnumerator rowEnumerator = worksheet.Cells.Rows.GetEnumerator(false, true);
        while (rowEnumerator.MoveNext())
        {
            Row row = (Row)rowEnumerator.Current;
            Console.WriteLine($"[Sync] Row {row.Index} visited.");
        }

        // Save the workbook (no modifications made in this example)
        workbook.Save("Processed.xlsx");
    }
}