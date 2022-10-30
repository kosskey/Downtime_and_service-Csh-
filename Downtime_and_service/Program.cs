global using Excel = Microsoft.Office.Interop.Excel;

namespace Downtime_and_service;

class Program
{
    static void Main(string[] args)
    {
        var config = Config.Config.Get_Config();

        var excel = File.File.Start_Excel();

        //файл отчета
        var report_config = new File.File("report", config);
        var excelWorkBook_report = report_config.Open_file(excel);

            var excelWorkSheet_rating = report_config.Activate_sheet(excelWorkBook_report, "Источник Рейтин");
            excelWorkSheet_rating.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            excel.Run("Выкладки");

            




        var operators_CA_config = new File.File("operators_CA", config);
        var excelWorkBook_operators_CA = operators_CA_config.Open_file(excel);

        var operators_CC_config = new File.File("operators_CC", config);
        var excelWorkBook_operators_CC = operators_CC_config.Open_file(excel);

        var technician_config = new File.File("technician", config);
        var excelWorkBook_technician = technician_config.Open_file(excel);

        var revenue_config = new File.File("revenue", config);
        var excelWorkBook_revenue = revenue_config.Open_file(excel);

        var amount_config = new File.File("amount", config);
        var excelWorkBook_amount = amount_config.Open_file(excel);

        var not_connection_config = new File.File("not_connection", config);
        var excelWorkBook_not_connection = not_connection_config.Open_file(excel);

        var not_work_config = new File.File("not_work", config);
        var excelWorkBook_not_work = not_work_config.Open_file(excel);


        var rating_config = new File.File("rating", config);
        var excelWorkBook_rating = rating_config.Open_file(excel);






        excelWorkBook_report.Close();
        excelWorkBook_operators_CA.Close();
        excelWorkBook_operators_CC.Close();
        excelWorkBook_technician.Close();
        excelWorkBook_revenue.Close();
        excelWorkBook_amount.Close();
        excelWorkBook_not_connection.Close();
        excelWorkBook_not_work.Close();
        excelWorkBook_rating.Close();
        
        
        
        /*
        
        

        Console.WriteLine("Введите дату отчета в формате ГГГГ.ММ.ДД");
        string? d_reverse = Console.ReadLine();
        
        var date = new Date(d_reverse!);
        
        var date = new Dictionary<string, string>()
        {
            ["day"] = d_reverse!.Substring(8, 2),
            ["month"] = d_reverse.Substring(5, 2),
            ["year"] = d_reverse.Substring(0, 4)
        };
        
        string d_full = date["day"] + "." + date["month"] + "." + date["year"];
        string d_briefly = $"{date["day"]}.{date["month"]}";
        

        Excel.Workbook? ExcelWorkBook_report = null;  //файл Отчет
        var ExcelWorkBook_sources = new Dictionary<string, Excel.Workbook>(); //файлы Исходников
        Excel.Workbook? ExcelWorkBook_rating = null; //файл Рейтингов

        while (true) {
            Console.WriteLine("Выберите функцию:");
            Console.WriteLine("1. Открыть файлы отчета");
            Console.WriteLine("2. Копировать сведения");
            Console.WriteLine("3. Сохраниить и закрыть все файлы");
            Console.WriteLine("4. Завершить скрипт");

            string? v = Console.ReadLine();

            if (v == "1")
            {
                ExcelWorkBook_report = FileAction.FuncOpen1(ExcelObj, date);
                FileAction.FuncOpen2(ExcelObj, date, d_briefly, ExcelWorkBook_sources);
                ExcelWorkBook_rating = FileAction.FuncOpen3(ExcelObj, date, d_briefly, d_full);
            }
            else if(v == "2")
            {
                FileAction.FuncCopy(ExcelObj, date, d_briefly, ExcelWorkBook_report!, ExcelWorkBook_sources, ExcelWorkBook_rating!);
            }
            else if(v == "3")
            {
                FileAction.FuncClose(ExcelObj, date, d_full, ExcelWorkBook_report!, ExcelWorkBook_sources, ExcelWorkBook_rating!);
            }
            else if(v == "4")
            {
                //System.Diagnostics.Process ExcelProcess = new System.Diagnostics.Process();

                var Ex = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                Ex[0].Kill();
                break;
            }
        }
        */
    }
}