using System;

namespace Downtime_and_service;

class ClassWorkbook
{
    public static void Report_create(Dictionary<string, string> config, Excel.Application excel, File.ClassFile report_config, Excel.Workbook excelWorkBook_report)
    {
        var excelWorkSheet_rating = report_config.Activate_sheet(excelWorkBook_report, "Источник Рейтинг");
        excelWorkSheet_rating.Visible = Excel.XlSheetVisibility.xlSheetVisible;
        excel.Run("Выкладки");

        var excelWorkSheet_install = report_config.Activate_sheet(excelWorkBook_report, "Установочные");
        string? amount_line_old_start = (string)excelWorkSheet_install.Range["B2"].Text;
        string? amount_line_old_end = Convert.ToString(Int32.Parse(amount_line_old_start) - 13);
        string? amount_line_next_start = Convert.ToString(Convert.ToInt32(amount_line_old_start) + 1);
        string? amount_line_next_end = Convert.ToString(Convert.ToInt32(amount_line_old_start) + 14);
            
        excelWorkSheet_rating.Range["A" + amount_line_old_start + ":AC" + amount_line_old_end].Copy();
        excelWorkSheet_rating.Paste(excelWorkSheet_rating.Range["A" + amount_line_next_start]);
        //DateTime date1 = new DateTime(Int32.Parse(config["year"]), Int32.Parse(config["month"]), Int32.Parse(config["day"]));
        //DateTime date2 = DateTime.Parse(config["date_current_report"]);
        DateTime date = Convert.ToDateTime(config["date_current_report"]);
        excelWorkSheet_rating.Range["B" + amount_line_next_start + ":B" + amount_line_next_end].Value = date;
    }

    public static void Report_copy()
    {

    }

    public static void Sources_create(Dictionary<string, string> config, File.ClassFile sources_config, Excel.Workbook excelWorkBook_sources)
    {
        int last_sheet = excelWorkBook_sources.Worksheets.Count;

        int checker = 0;
        foreach (Excel.Worksheet sheet_existing in excelWorkBook_sources.Sheets)
        {
            if (sheet_existing.Name == config["date_current_report"].Substring(0, 5))
            {
                checker = 1;
            }
        }
        if (checker == 1)
        {
            var ExcelWorkSheet_sheet1 = (Excel.Worksheet)excelWorkBook_sources.Worksheets.Add(System.Reflection.Missing.Value, excelWorkBook_sources.Worksheets[last_sheet]);
            ExcelWorkSheet_sheet1.Name = "Лист1";
        }
        else
        {
            var ExcelWorkSheet_brefly = (Excel.Worksheet)excelWorkBook_sources.Worksheets.Add(System.Reflection.Missing.Value, excelWorkBook_sources.Worksheets[last_sheet]);
            ExcelWorkSheet_brefly.Name = config["date_current_report"].Substring(0, 5);
        }

        var ExcelWorkSheet_install = sources_config.Activate_sheet(excelWorkBook_sources, "Установочные");
        ExcelWorkSheet_install.Range["B1"].Value = config["date_current_report"].Substring(0, 5);
    }

    public static void Sources_copy()
    {

    }

    public static void Rating_create(Dictionary<string, string> config, Excel.Application excel, File.ClassFile rating_config, Excel.Workbook excelWorkBook_rating)
    {
        int last_sheet = excelWorkBook_rating.Worksheets.Count;
        var ExcelWorkSheet_new_list = (Excel.Worksheet)excelWorkBook_rating.Worksheets.Add(System.Reflection.Missing.Value, excelWorkBook_rating.Worksheets[last_sheet]);
        ExcelWorkSheet_new_list.Name = config["date_current_report"].Substring(0, 5);
        ExcelWorkSheet_new_list.Range["A1"].Value = "Рейтинг подразделений по 4 показателям за " + config["date_current_report"];
    }

    public static void Rating_copy()
    {

    }
}