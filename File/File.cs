namespace File;

public class ClassFile
{
    //private readonly Excel.Application ExcelObj = Start_Excel();
    //private readonly Excel.Application ExcelObj;

    private readonly string name_eng;
    private readonly string name_rus;
    private readonly string link;

    public ClassFile(string name_eng, Dictionary<string, string> config)
    {
        var date = new Date.ClassDate(config["date_current_report"]);

        this.name_eng = "ExcelWorkBook_" + name_eng;

        if (name_eng == "report")
        {
            name_rus = config[this.name_eng] + config["date_previous_report"] + ".xlsm";
        }
        else if (name_eng == "rating")
        {
            name_rus = config[this.name_eng] + date.month_name[date.month] + " " + date.yaer + ".xlsx";
        }
        else
        {
            name_rus = config["folder"] + date.month_name[date.month] + " " + date.yaer + "\\" + config[this.name_eng] + ".xlsx";
        }

        link = config["path_directory"] + date.yaer + "\\" + date.month + ". " + date.month_name[date.month] + "\\" + name_rus;
    }

    public static Excel.Application Start_Excel()
    {
        var excelObj = new Excel.Application();
        excelObj.Visible = true;
        excelObj.WindowState = Excel.XlWindowState.xlMaximized;
        return excelObj;
    }

    public Excel.Workbook Open_file(Excel.Application excel)
    {
        var workbook = excel.Workbooks.Open(link);

        return workbook;
    }

    public Excel.Worksheet Activate_sheet(Excel.Workbook workbook, string name_sheet)
    {
        var worksheet = (Excel.Worksheet)workbook.Sheets.Item[name_sheet];
        
        return worksheet;
    }

    public static void Close_file(Excel.Workbook workbook)
    {
        if (workbook != null)
        {
            workbook.Close();
        }
    }

    /*
    public static void FuncClose(Excel.Application ExcelObj, Dictionary<string, string> date, string d_full, Excel.Workbook ExcelWorkBook_report, Dictionary<string, Excel.Workbook> ExcelWorkBook_sources, Excel.Workbook ExcelWorkBook_rating)
        {
            FilePath Path = new FilePath(date);

            var ExcelWorkSheet = (Excel.Worksheet)ExcelWorkBook_report.Sheets.Item["Источник Рейтинг"];
            ExcelWorkSheet.Activate();
            ExcelObj.Run("Выкладки");
            ExcelWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            //Format-List -Property Name, Index -InputObject $ExcelWorkBook_report.Sheets.Item("Рейтинг")
            var ExcelWorkSheet_active = (Excel.Worksheet)ExcelWorkBook_report.Sheets[1];
            ExcelWorkSheet_active.Activate();
            ExcelObj.DisplayAlerts = false;
            ExcelWorkBook_report.SaveAs(Path.path_directory + "Отчет по простоям и сервису_" + d_full + ".xlsm", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
            ExcelWorkBook_report.Close();

            //foreach (key in ExcelWorkBook_sources.Keys) {ExcelWorkBook_sources[key].Save() ExcelWorkBook_sources[key].Close()}

            foreach (string node in Path.source_name_eng)
            {
                ExcelWorkBook_sources[node].Save();
                ExcelWorkBook_sources[node].Close();
            };

            ExcelWorkBook_rating.Save();
            ExcelWorkBook_rating.Close();
        }
    */
}
