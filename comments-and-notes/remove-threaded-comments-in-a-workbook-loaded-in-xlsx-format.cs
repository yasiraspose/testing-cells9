using System;
using Aspose.Cells;

public class Program
{
    public static void Main()
    {
        RemoveThreadedCommentsDemo.Run();
    }
}

public class RemoveThreadedCommentsDemo
{
    public static void Run()
    {
        Workbook workbook = new Workbook("input.xlsx");
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments();
        }
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}