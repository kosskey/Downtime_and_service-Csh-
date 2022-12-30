namespace Files;

public class Path
{
    private readonly string name_eng;
    private readonly string name_rus;
    private readonly string file_name;
    public readonly string path_file;
    public readonly string path_report_for_save;
    
    public Path(string name_eng, Dictionary<string, string> config, Date.ClassDate date, string filename_extension = null!)
    {
        this.name_eng = "ExcelWorkBook_" + name_eng;
        name_rus = config[this.name_eng];

        if (name_eng == "report")
        {
            file_name = config[this.name_eng] + config["date_previous_report"] + ".xlsm";
        }
        else if (name_eng == "rating")
        {
            file_name = config[this.name_eng] + date.month_name[date.month] + " " + date.yaer + ".txt";
        }
        else if (name_eng == "not_connection" || name_eng == "not_work")
        {
            file_name = config["folder"] + date.month_name[date.month] + " " + date.yaer + "\\" + config[this.name_eng] + ".txt";
        }
        else
        {
            file_name = config["folder"] + date.month_name[date.month] + " " + date.yaer + "\\" + config[this.name_eng] + "." + filename_extension;
        }

        path_file = config["path_directory"] + date.yaer + "\\" + date.month + ". " + date.month_name[date.month] + "\\" + file_name;
        path_report_for_save = config["path_directory"] + date.yaer + "\\" + date.month + ". " + date.month_name[date.month] + "\\" + "Отчет по простоям и сервису_" + config["date_current_report"] + ".xlsm";
    }
}

public class FileXLSX
{
    public static Excel.Application Start_Excel()
    {
        var excelObj = new Excel.Application();
        excelObj.Visible = true;
        excelObj.WindowState = Excel.XlWindowState.xlMaximized;
        return excelObj;
    }

    public static Excel.Workbook Open_file(Excel.Application excel, string path_open)
    {
        var workbook = excel.Workbooks.Open(path_open);

        return workbook;
    }

    public static Excel.Worksheet Activate_sheet(Excel.Workbook workbook, string name_sheet)
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
    /*
    public static void Convert_ANSI_UTF8(string name_eng, Dictionary<string, string> config, Files.Path save_link)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        StreamReader sr = new StreamReader(save_link + config["ExcelWorkBook_" + name_eng] + ".csv", System.Text.Encoding.GetEncoding("Windows-1251"));
        StreamWriter sw = new StreamWriter(save_link + config["ExcelWorkBook_" + name_eng] + "2.csv", true, System.Text.Encoding.UTF8);

        var line = sr.ReadLine();
        while (line != null)
        {
            line = sr.ReadLine();
            sw.WriteLine(line);
        }

        sr.Close();
        sw.Close();
    }
    */
    
    public static void Open_CSV(string name_eng, Dictionary<string, string> config, string path_file)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        StreamReader sr = new StreamReader(path_file, System.Text.Encoding.GetEncoding("Windows-1251"));
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
                Database.ClassDatabase.Insert(name_eng, slist, null!, config["date_current_report"]);
                line = sr.ReadLine()!;
            }
            line = sr.ReadLine()!;
        }
        sr.Close();
    }
}

public class FileTXT
{
    public static void Open_TXT(string name_eng, Dictionary<string, string> config, string path_file)
    {
        StreamReader sr = new StreamReader(path_file);
        String line = sr.ReadLine()!;

        while (line != null)
        {
            if (line != "" )
            {
                string[] slist = line!.Split(' ');
                string slist2 = sr.ReadLine()!;

                Database.ClassDatabase.Insert(name_eng, slist, slist2, config["date_current_report"]);
            }
            line = sr.ReadLine()!;
        }
        line = sr.ReadLine()!;

        sr.Close();
    }
}
