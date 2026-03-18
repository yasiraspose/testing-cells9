using System;
using Aspose.Cells;

namespace LightCellsExample
{
    // Custom LightCellsDataProvider that streams a large dataset to the workbook
    public class LargeDataProvider : LightCellsDataProvider
    {
        private const int MaxRows = 100_000;   // Example large number of rows
        private const int MaxCols = 10;        // Example number of columns per row

        private int _currentRow = -1;
        private int _currentCol = -1;
        private bool _processSheet = false;

        // Called once for each worksheet being saved
        public bool StartSheet(int sheetIndex)
        {
            // Process only the first worksheet (index 0)
            _processSheet = sheetIndex == 0;
            return _processSheet;
        }

        // Returns the next row index to be saved, or -1 when done
        public int NextRow()
        {
            if (!_processSheet) return -1;

            _currentRow++;
            _currentCol = -1; // reset column index for the new row

            return _currentRow < MaxRows ? _currentRow : -1;
        }

        // Allows optional row-level configuration
        public void StartRow(Row row)
        {
            // Example: set a fixed row height
            row.Height = 15;
        }

        // Returns the next column index for the current row, or -1 when the row is finished
        public int NextCell()
        {
            _currentCol++;
            return _currentCol < MaxCols ? _currentCol : -1;
        }

        // Sets the value for the current cell
        public void StartCell(Cell cell)
        {
            // Example value: "R{row}C{col}"
            cell.PutValue($"R{_currentRow}C{_currentCol}");
        }

        // Determines whether string values should be gathered into a global pool
        public bool IsGatherString()
        {
            // For this example we do not need a global string pool
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Configure OoxmlSaveOptions with the custom LightCellsDataProvider
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new LargeDataProvider()
            };

            // Save the workbook using the LightCells mode (uses the Save(string, SaveOptions) rule)
            workbook.Save("LargeDataOutput.xlsx", saveOptions);

            Console.WriteLine("Large workbook saved successfully using LightCells API.");
        }
    }
}