using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the XLSX workbook with Russian culture settings
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CultureInfo = new CultureInfo("ru-RU");
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Apply custom globalization settings for Russian boolean and error values
        workbook.Settings.GlobalizationSettings = new RussianGlobalizationSettings();

        // Optional: display localized values of the first row for verification
        Cells cells = workbook.Worksheets[0].Cells;
        for (int col = 0; col <= cells.MaxDataColumn; col++)
        {
            Console.WriteLine($"Cell[0,{col}]: {cells[0, col].StringValue}");
        }

        // Save the localized workbook
        workbook.Save("output.xlsx");
    }

    // Custom globalization settings that translate booleans and error strings to Russian
    class RussianGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool value)
        {
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string error)
        {
            switch (error)
            {
                case "#NAME?":   return "#ИМЯ?";
                case "#DIV/0!":  return "#ДЕЛ/0!";
                case "#REF!":    return "#ССЫЛКА!";
                case "#VALUE!":  return "#ЗНАЧ!";
                case "#N/A":     return "#Н/Д";
                case "#NUM!":    return "#ЧИСЛО!";
                case "#NULL!":   return "#ПУСТО!";
                default:         return base.GetErrorValueString(error);
            }
        }
    }
}