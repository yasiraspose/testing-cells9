using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

class RemoveThreadedCommentsDemo
{
    static void Main()
    {
        string inputPath = "InputWithThreadedComments.xlsx";
        Workbook workbook;

        if (File.Exists(inputPath))
        {
            workbook = new Workbook(inputPath);
        }
        else
        {
            workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // Add a regular comment (optional)
            int commentIdx = ws.Comments.Add("A1");
            ws.Comments[commentIdx].Note = "Sample comment";

            // Add a threaded comment (optional) using reflection if supported
            PropertyInfo tcProp = typeof(Worksheet).GetProperty("ThreadedComments");
            if (tcProp != null)
            {
                object threadedComments = tcProp.GetValue(ws);
                MethodInfo addMethod = threadedComments.GetType().GetMethod("Add", new[] { typeof(string) });
                if (addMethod != null)
                {
                    object idxObj = addMethod.Invoke(threadedComments, new object[] { "B2" });
                    int idx = Convert.ToInt32(idxObj);
                    PropertyInfo itemProp = threadedComments.GetType().GetProperty("Item");
                    object tc = itemProp.GetValue(threadedComments, new object[] { idx });
                    PropertyInfo commentProp = tc.GetType().GetProperty("Comment");
                    commentProp.SetValue(tc, "Sample threaded comment");
                }
            }
        }

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments();

            // Remove threaded comments if the API is available
            PropertyInfo tcProp = typeof(Worksheet).GetProperty("ThreadedComments");
            if (tcProp != null)
            {
                object threadedComments = tcProp.GetValue(sheet);
                MethodInfo clearMethod = threadedComments.GetType().GetMethod("Clear");
                clearMethod?.Invoke(threadedComments, null);
            }
        }

        workbook.Save("OutputWithoutThreadedComments.xlsx");
    }
}