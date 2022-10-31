namespace Downtime_and_service;

class ClassProgram
{
    static void Main(string[] args)
    {
        var config = Config.ClassConfig.Get_Config();

        Console.WriteLine("Дата текущего отчета: " + config["date_current_report"]);
        Console.WriteLine("Дата предыдущего отчета: " + config["date_previous_report"]);
        
        while (true) {
            Console.WriteLine("Выберите функцию:");
            Console.WriteLine("1. Запустить Excel, открыть все файлы отчетов, создать новые выкладки");
            Console.WriteLine("2. Копировать сведения");
            Console.WriteLine("3. Сохраниить и закрыть все файлы отчетов");
            Console.WriteLine("4. Зактрыть Excel, завершить скрипт");

            string? v = Console.ReadLine();

            Excel.Workbook? excelWorkBook_report = null;
            
            if (v == "1")
            {
                var excel = File.ClassFile.Start_Excel();
                
                var report_config = new File.ClassFile("report", config);
                excelWorkBook_report = report_config.Open_file(excel);
                ClassWorkbook.Report_create(config, excel, report_config, excelWorkBook_report);
/*
                var operators_CA_config = new File.ClassFile("operators_CA", config);
                var excelWorkBook_operators_CA = operators_CA_config.Open_file(excel);

                var operators_CC_config = new File.ClassFile("operators_CC", config);
                var excelWorkBook_operators_CC = operators_CC_config.Open_file(excel);

                var technician_config = new File.ClassFile("technician", config);
                var excelWorkBook_technician = technician_config.Open_file(excel);

                var revenue_config = new File.ClassFile("revenue", config);
                var excelWorkBook_revenue = revenue_config.Open_file(excel);

                var amount_config = new File.ClassFile("amount", config);
                var excelWorkBook_amount = amount_config.Open_file(excel);

                var not_connection_config = new File.ClassFile("not_connection", config);
                var excelWorkBook_not_connection = not_connection_config.Open_file(excel);

                var not_work_config = new File.ClassFile("not_work", config);
                var excelWorkBook_not_work = not_work_config.Open_file(excel);

                var rating_config = new File.ClassFile("rating", config);
                var excelWorkBook_rating = rating_config.Open_file(excel);
*/
            }
            else if(v == "2")
            {
                //FileAction.FuncCopy(ExcelObj, date, d_briefly, ExcelWorkBook_report!, ExcelWorkBook_sources, //ExcelWorkBook_rating!);
            }
            else if(v == "3")
            {
                File.ClassFile.Close_file(excelWorkBook_report!);
                
                //FileAction.FuncClose(ExcelObj, date, d_full, ExcelWorkBook_report!, ExcelWorkBook_sources, ExcelWorkBook_rating!);
            }
            else if(v == "4")
            {
                //System.Diagnostics.Process ExcelProcess = new System.Diagnostics.Process();

                //var Ex = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                //Ex[0].Kill();
                //break;
            }
        }
    }
}