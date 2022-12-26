namespace TypeFile;

public class FileXLSX
{
    //private readonly Excel.Application ExcelObj = Start_Excel();
    //private readonly Excel.Application ExcelObj;

    private readonly string name_eng;
    public readonly string name_rus;
    private readonly string link_file;
    private readonly string full_link;
    public readonly string save_link;

    public FileXLSX(string name_eng, Dictionary<string, string> config, Date.ClassDate date)
    {
        this.name_eng = "ExcelWorkBook_" + name_eng;
        name_rus = config[this.name_eng];

        if (name_eng == "report")
        {
            link_file = config[this.name_eng] + config["date_previous_report"] + ".xlsm";
        }
        else if (name_eng == "rating")
        {
            link_file = config[this.name_eng] + date.month_name[date.month] + " " + date.yaer + ".xlsx";
        }
        else
        {
            link_file = config["folder"] + date.month_name[date.month] + " " + date.yaer + "\\" + config[this.name_eng] + ".xlsx";
        }

        full_link = config["path_directory"] + date.yaer + "\\" + date.month + ". " + date.month_name[date.month] + "\\" + link_file;
        save_link = config["path_directory"] + date.yaer + "\\" + date.month + ". " + date.month_name[date.month] + "\\" + "Отчет по простоям и сервису_" + config["date_current_report"] + ".xlsm";
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
        var workbook = excel.Workbooks.Open(full_link);

        return workbook;
    }

    public Excel.Worksheet Activate_sheet(Excel.Workbook workbook, string name_sheet)
    {
        var worksheet = (Excel.Worksheet)workbook.Sheets.Item[name_sheet];
        
        return worksheet;
    }

    public Excel.Range Activate_range(Excel.Worksheet worksheet, int row, int column)
    {
        var workrange = (Excel.Range)worksheet.Cells.Item[row, column];
        
        return workrange;
    }

    public static void Close_file(Excel.Workbook workbook)
    {
        if (workbook != null)
        {
            workbook.Close();
        }
    }
}

public class FileCSV
{
    public static void Convert_XLSX_to_CSV(Excel.Application excel)
    {
        //var workbook = excel.Workbooks.Open();
    }
    
    public static void Open_CSV(string name_eng, Dictionary<string, string> config)
    {
        StreamReader sr = new StreamReader(Directory.GetCurrentDirectory() + "\\Работа\\2. Отчеты\\1. Ежедневный\\4. Простои и сервис\\2022\\09. Сентябрь\\Исходники из 1С_Сентябрь 2022" + "\\" + config["ExcelWorkBook_" + name_eng] + ".csv");
        String line = sr.ReadLine()!;

        while (line != null)
        {
            line = sr.ReadLine()!;
            while (line != "" & line != null)
            {
                string[] slist = line!.Split(';');
                for (int i = 0; i < slist.Length; i++)
                {
                    if (slist[i] == "")
                    {
                        slist[i] = null!;
                    }
                }
                Database.ClassDatabase.Insert(name_eng, slist);
                line = sr.ReadLine()!;
            }
            line = sr.ReadLine()!;
        }
        sr.Close();
    }
}

public class FileTXT
{

}
