using System;
using Aspose.Cells;

class AutoFilterDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Category");
        worksheet.Cells["C1"].PutValue("Price");

        worksheet.Cells["A2"].PutValue("Laptop");
        worksheet.Cells["B2"].PutValue("Electronics");
        worksheet.Cells["C2"].PutValue(1200);

        worksheet.Cells["A3"].PutValue("Shirt");
        worksheet.Cells["B3"].PutValue("Clothing");
        worksheet.Cells["C3"].PutValue(45);

        worksheet.Cells["A4"].PutValue("Phone");
        worksheet.Cells["B4"].PutValue("Electronics");
        worksheet.Cells["C4"].PutValue(800);

        worksheet.Cells["A5"].PutValue("Book");
        worksheet.Cells["B5"].PutValue("Books");
        worksheet.Cells["C5"].PutValue(20);

        // Apply an AutoFilter to the range that includes the header and data rows
        worksheet.AutoFilter.Range = "A1:C5";

        // Filter the 'Category' column (field index 1) to show only "Electronics"
        worksheet.AutoFilter.Filter(1, "Electronics");
        worksheet.AutoFilter.Refresh();

        // Save the workbook with the applied filter
        workbook.Save("AutoFilterDemo.xlsx");
    }
}